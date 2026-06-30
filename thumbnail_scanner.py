import os
import sys
import subprocess
import json
import concurrent.futures
import shutil
import re
import hashlib
import tempfile
try:
    import win32com.client
except ImportError:
    win32com = None
from datetime import datetime, timezone

VIDEO_EXTENSIONS = ('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.webm')
IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.webp', '.JPG', '.JPEG', '.PNG', '.WEBP')

def get_projects(base_path):
    return [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d)) and not d.startswith('.') and d.lower() != "originals"]

def find_videos(project_path):
    videos = []
    skip_dirs = {'.git', '__pycache__', 'thumbnails', 'edit thumbnails', '$recycle.bin', 'system volume information', 'originals'}
    for root, dirs, files in os.walk(project_path):
        dirs[:] = [d for d in dirs if d.lower() not in skip_dirs]
        for file in files:
            if file.lower().endswith(VIDEO_EXTENSIONS):
                videos.append(os.path.join(root, file))
    return videos

def check_thumbnails_optimized(video_path, project_path, thumb_files, edit_files):
    video_name = os.path.splitext(os.path.basename(video_path))[0]
    
    main_found_file = None
    for ext in IMAGE_EXTENSIONS:
        f = video_name + ext
        if f in thumb_files:
            main_found_file = f
            break
            
    edit_found_files = []
    edit_indices_found = []
    for i in range(1, 11):
        for ext in IMAGE_EXTENSIONS:
            f = f"{video_name}_{i}{ext}"
            if f in edit_files:
                edit_found_files.append(f)
                edit_indices_found.append(i)
                break
                
    return main_found_file, edit_indices_found, edit_found_files

def get_video_info(video_path):
    """Get total frames, fps, and dimensions (including DAR) using ffprobe."""
    cmd = [
        'ffprobe', '-v', 'error', '-select_streams', 'v:0',
        '-show_entries', 'stream=nb_frames,avg_frame_rate,width,height,display_aspect_ratio,sample_aspect_ratio',
        '-of', 'json', video_path
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        data = json.loads(result.stdout)
        streams = data.get('streams', [])
        if not streams: return 0, 25.0, 0, 0, 1.0
        stream = streams[0]
        
        fps = 25.0
        avg_frame_rate = stream.get('avg_frame_rate', '25/1')
        if '/' in avg_frame_rate:
            parts = avg_frame_rate.split('/')
            if len(parts) == 2:
                try:
                    num, den = map(int, parts)
                    if den != 0: fps = num / den
                except ValueError: pass
        else:
            try: fps = float(avg_frame_rate)
            except ValueError: pass
            
        nb_frames = int(stream.get('nb_frames', 0))
        if nb_frames == 0:
            try:
                cmd_dur = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'json', video_path]
                res_dur = subprocess.run(cmd_dur, capture_output=True, text=True, check=True)
                dur = float(json.loads(res_dur.stdout).get('format', {}).get('duration', 0))
                nb_frames = int(dur * fps)
            except Exception: pass
            
        width = int(stream.get('width', 0))
        height = int(stream.get('height', 0))
        
        # Determine actual aspect ratio (DAR)
        dar_val = None
        
        # Priority 1: display_aspect_ratio
        dar_str = stream.get('display_aspect_ratio')
        if dar_str and dar_str not in ("0:1", "0/1"):
            if ':' in dar_str or '/' in dar_str:
                sep = ':' if ':' in dar_str else '/'
                try:
                    num, den = map(int, dar_str.split(sep))
                    if den != 0: dar_val = num / den
                except ValueError: pass
            else:
                try:
                    dar_val = float(dar_str)
                except ValueError: pass
        
        if dar_val is None or dar_val <= 0:
            sar_val = 1.0
            sar_str = stream.get('sample_aspect_ratio')
            if sar_str and sar_str != "0:1":
                if ':' in sar_str:
                    try:
                        num, den = map(int, sar_str.split(':'))
                        if den != 0: sar_val = num / den
                    except ValueError: pass
                else:
                    try:
                        sar_val = float(sar_str)
                    except ValueError: pass
            dar_val = (width / height) * sar_val if height != 0 else 1.0
        
        if dar_val <= 0: dar_val = 1.0
        return nb_frames, fps, width, height, dar_val
    except Exception:
        return 0, 25.0, 0, 0, 1.0

def get_md5(path):
    if not os.path.exists(path): return None
    hash_md5 = hashlib.md5()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest().upper()
    except Exception:
        return None

def get_image_dimensions(image_path):
    """Get width and height of an image using ffprobe."""
    cmd = [
        'ffprobe', '-v', 'error', '-select_streams', 'v:0',
        '-show_entries', 'stream=width,height',
        '-of', 'json', image_path
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        data = json.loads(result.stdout)
        stream = data.get('streams', [{}])[0]
        return int(stream.get('width', 0)), int(stream.get('height', 0))
    except Exception:
        return 0, 0

def generate_video_thumbnails(task):
    """Worker function using FFmpeg processes with dimension logic."""
    video_path, project_path, gen_main, missing_edits = task
    video_name = os.path.splitext(os.path.basename(video_path))[0]
    
    nb_frames, fps, v_width, v_height, v_ar = get_video_info(video_path)
    if nb_frames <= 0:
        return video_path, False

    thumb_dir = os.path.join(project_path, 'Thumbnails')
    edit_dir = os.path.join(project_path, 'Edit Thumbnails')
    
    # Pre-calculate target frames for all 11 slots (Main + 10 Edits)
    all_target_frames = [min(1, nb_frames - 1)] # Slot 0: Main
    start_idx = 1
    end_idx = max(1, nb_frames - 2)
    step = (end_idx - start_idx) / 9
    for i in range(10):
        all_target_frames.append(int(start_idx + i * step)) # Slots 1-10: Edits

    # Determine which slots need generation
    slots_to_generate = []
    if gen_main:
        slots_to_generate.append(0)
    for i in missing_edits:
        slots_to_generate.append(i)
    
    if not slots_to_generate:
        return video_path, True

    # Determine which group (Portrait, Square, Landscape) this video belongs to
    # New Target Dimensions (approx 50% of previous): 
    # Portrait (284x504 = 0.5635), Square (504x504 = 1.0), Landscape (896x504 = 1.777)
    diffs = {
        'portrait': (abs(v_ar - 0.5625), 284, 504),
        'square': (abs(v_ar - 1.0), 504, 504),
        'landscape': (abs(v_ar - 1.777), 896, 504)
    }
    group_name, (_, target_w, target_h) = min(diffs.items(), key=lambda x: x[1][0])

    # 10% Tolerance Logic
    # We want to scale the video dimensions to fit into target_w x target_h whilst preserving v_ar.
    if v_ar >= (target_w / target_h):
        rw, rh = target_w, int(target_w / v_ar)
    else:
        rh, rw = target_h, int(target_h * v_ar)

    w_diff = abs(rw - target_w) / target_w
    h_diff = abs(rh - target_h) / target_h
    use_padding = w_diff > 0.10 or h_diff > 0.10

    # Build filters
    filter_list = [f"scale={target_w}:{target_h}:force_original_aspect_ratio=decrease"]
    if use_padding:
        filter_list.append(f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2")
    
    filter_list.extend(["setsar=1", "format=yuvj420p"])

    success = True
    try:
        unique_frames = sorted(list(set(all_target_frames[s] for s in slots_to_generate)))
        select_str = " + ".join([f"eq(n,{idx})" for idx in unique_frames])
        
        filter_parts = [f"select='{select_str}'"] + filter_list + ["setpts=N/FRAME_RATE/TB"]
        filter_graph = ",".join(filter_parts)
            
        temp_pattern = os.path.join(project_path, f"tmp_{video_name}_%d.jpg")
        cmd = [
            'ffmpeg', '-y', '-threads', '1', '-i', video_path,
            '-vf', filter_graph, '-vsync', 'vfr', '-q:v', '2',
            temp_pattern
        ]
        
        if subprocess.run(cmd, capture_output=True).returncode != 0:
            success = False
        else:
            # Move/Rename
            os.makedirs(thumb_dir, exist_ok=True)
            os.makedirs(edit_dir, exist_ok=True)
            
            for slot in slots_to_generate:
                frame_idx = all_target_frames[slot]
                out_idx = unique_frames.index(frame_idx) + 1
                src = temp_pattern % out_idx
                if slot == 0:
                    dst = os.path.join(thumb_dir, f"{video_name}.jpg")
                else:
                    dst = os.path.join(edit_dir, f"{video_name}_{slot}.jpg")
                if os.path.exists(src):
                    shutil.copy2(src, dst)
            
            # Cleanup
            for i in range(1, len(unique_frames) + 1):
                p = temp_pattern % i
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
        
        # Fallback for slot 10 if it failed to generate
        if 10 in missing_edits:
            slot_10_path = os.path.join(edit_dir, f"{video_name}_10.jpg")
            if not os.path.exists(slot_10_path):
                fallback_filter_str = ",".join(filter_list)
                # Using -sseof -1 allows seeking to 1 second before end
                cmd_fallback = [
                    'ffmpeg', '-y', '-sseof', '-1', '-i', video_path,
                    '-vf', fallback_filter_str, '-update', '1', '-frames:v', '1', '-q:v', '2',
                    slot_10_path
                ]
                if subprocess.run(cmd_fallback, capture_output=True).returncode != 0:
                    # If -sseof -1 fails (e.g. video < 1s), try without seeking
                    cmd_fallback_no_seek = [
                        'ffmpeg', '-y', '-i', video_path,
                        '-vf', fallback_filter_str, '-frames:v', '1', '-q:v', '2',
                        slot_10_path
                    ]
                    subprocess.run(cmd_fallback_no_seek, capture_output=True)

        return video_path, success
    except Exception:
        return video_path, False

def is_shortcut(path):
    if os.path.islink(path): return True
    if path.lower().endswith('.lnk'): return True
    return False

_wscript_shell = None
def get_wscript_shell():
    global _wscript_shell
    if _wscript_shell is None and win32com:
        try:
            # We use Dispatch because it's persistent and avoids launching a process
            _wscript_shell = win32com.client.Dispatch("WScript.Shell")
        except Exception:
            pass
    return _wscript_shell

def get_shortcut_target(path):
    if os.path.islink(path):
        return os.readlink(path)
    if path.lower().endswith('.lnk'):
        if os.name == 'nt':
            shell = get_wscript_shell()
            if shell:
                try:
                    return shell.CreateShortcut(path).TargetPath
                except Exception:
                    pass
            # Fallback to PowerShell if COM fails or pywin32 is missing
            escaped_path = path.replace("'", "''")
            cmd = ['powershell', '-Command', f"(New-Object -ComObject WScript.Shell).CreateShortcut('{escaped_path}').TargetPath"]
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, check=True)
                return result.stdout.strip()
            except Exception: return None
    return None

def get_shortcut_targets_bulk(paths):
    if not paths: return {}
    results = {}
    if os.name == 'nt':
        shell = get_wscript_shell()
        if shell:
            for p in paths:
                try:
                    target = shell.CreateShortcut(p).TargetPath
                    if target:
                        results[p] = target
                except Exception:
                    pass
            return results

        # Use a temporary file to pass paths to PowerShell to avoid command line length limits
        # Use utf-8-sig to ensure PowerShell's Get-Content correctly handles the BOM
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8-sig') as tmp:
            for p in paths:
                tmp.write(p + '\n')
            tmp_path = tmp.name
        
        try:
            escaped_tmp_path = tmp_path.replace("'", "''")
            # Refactored for speed: using .NET ReadLines and a fast foreach loop
            ps_script = (
                "$s = New-Object -ComObject WScript.Shell; "
                f"[System.IO.File]::ReadLines('{escaped_tmp_path}') | ForEach-Object {{ "
                "try { $s.CreateShortcut($_).TargetPath } catch { '' } "
                "}"
            )
            cmd = ['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', ps_script]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            targets = result.stdout.splitlines()
            for p, t in zip(paths, targets):
                if t and t.strip():
                    results[p] = t.strip()
        except Exception as e:
            print(f"Error in bulk shortcut resolution: {e}")
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass
    else:
        for p in paths:
            if os.path.islink(p):
                try:
                    results[p] = os.readlink(p)
                except Exception:
                    pass
    return results

def update_shortcut(path, new_target):
    try:
        if os.path.islink(path):
            os.remove(path)
            os.symlink(new_target, path)
            return True
        if path.lower().endswith('.lnk') and os.name == 'nt':
            escaped_path = path.replace("'", "''")
            escaped_target = new_target.replace("'", "''")
            cmd = ['powershell', '-Command', f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{escaped_path}');$s.TargetPath='{escaped_target}';$s.Save()"]
            subprocess.run(cmd, check=True)
            return True
    except Exception as e:
        print(f"Error updating shortcut {path}: {e}")
    return False

def update_root_misc_shortcuts(base_path, targets_map, verbose=False):
    misc_path = os.path.join(base_path, 'misc.txt')
    existing_data = {}
    if os.path.exists(misc_path):
        try:
            with open(misc_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line: continue
                    parts = line.split('|')
                    if parts:
                        existing_data[parts[0]] = parts[1:]
        except Exception as e:
            if verbose: print(f"Error reading misc.txt: {e}")

    current_shortcut_names = set()
    for lnk_full_path, target in targets_map.items():
        shortcut_name = os.path.basename(lnk_full_path)
        current_shortcut_names.add(shortcut_name)
        
        try:
            mtime = os.path.getmtime(lnk_full_path)
            dt = datetime.fromtimestamp(mtime, timezone.utc).replace(tzinfo=None)
            iso_mtime = dt.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
            
            # Determine project folder and video name
            try:
                rel_to_base = os.path.relpath(target, base_path)
                if not rel_to_base.startswith('..'):
                    path_parts = rel_to_base.split(os.sep)
                    project_folder = path_parts[0]
                else:
                    project_folder = os.path.basename(os.path.dirname(target))
            except Exception:
                project_folder = os.path.basename(os.path.dirname(target))
                
            video_name = os.path.basename(target)
            
            sc_project = f"sc_project:{project_folder}"
            sc_video = f"sc_video:{video_name}"
            sc_mtime = f"sc_mtime:{iso_mtime}"
            
            if shortcut_name not in existing_data:
                existing_data[shortcut_name] = []
            
            # Filter out old shortcut-related parts
            new_parts = [p for p in existing_data[shortcut_name] 
                         if not (p.startswith('sc_project:') or 
                                 p.startswith('sc_video:') or 
                                 p.startswith('sc_mtime:'))]
            new_parts.extend([sc_project, sc_video, sc_mtime])
            existing_data[shortcut_name] = new_parts
        except Exception as e:
            if verbose: print(f"Error processing shortcut {shortcut_name}: {e}")

    # Cleanup: remove shortcut data for shortcuts that no longer exist in the root sc folder
    keys_to_delete = []
    for key in list(existing_data.keys()):
        has_sc_data = any(p.startswith('sc_project:') for p in existing_data[key])
        if has_sc_data and key not in current_shortcut_names:
            existing_data[key] = [p for p in existing_data[key] 
                                 if not (p.startswith('sc_project:') or 
                                         p.startswith('sc_video:') or 
                                         p.startswith('sc_mtime:'))]
            if not existing_data[key]:
                keys_to_delete.append(key)

    for key in keys_to_delete:
        del existing_data[key]

    if existing_data:
        if verbose:
            print(f"  - root: updating misc.txt with {len(current_shortcut_names)} shortcuts")
        with open(misc_path, 'w', encoding='utf-8') as f:
            for key in sorted(existing_data.keys()):
                line_parts = [key] + existing_data[key]
                f.write('|'.join(line_parts) + '\n')
    elif os.path.exists(misc_path):
        try:
            os.remove(misc_path)
        except Exception: pass

def update_sc_date(verbose=False):
    print("\nUpdating scdate.txt files...")
    base_path = os.path.abspath(os.getcwd())
    root_sc = os.path.join(base_path, 'sc')
    cached = []
    skip_dirs = {'.git', '__pycache__', 'thumbnails', 'edit thumbnails', '$recycle.bin', 'system volume information', 'originals'}
    
    # 1. Resolve root-level shortcuts first
    if os.path.exists(root_sc):
        print("Processing root sc shortcuts...")
        lnk_paths = [os.path.join(root_sc, f) for f in os.listdir(root_sc) if f.lower().endswith('.lnk')]
        if lnk_paths:
            targets_map = get_shortcut_targets_bulk(lnk_paths)
            for p, target in targets_map.items():
                try:
                    mtime = os.path.getmtime(p)
                    # Normalize target path for robust matching
                    norm_target = os.path.abspath(target)
                    cached.append({'target': norm_target, 'date': datetime.fromtimestamp(mtime, timezone.utc).replace(tzinfo=None)})
                except Exception: continue

    # 2. Identify all target directories
    print("Identifying project directories...")
    target_dirs = {base_path}
    for root, dirs, files in os.walk(base_path):
        dirs[:] = [d for d in dirs if d.lower() not in skip_dirs]
        # Any directory with an 'sc' subfolder is a project
        if 'sc' in dirs:
            target_dirs.add(os.path.abspath(root))
        # Any immediate subdirectory of base_path (excluding specials) is a project
        if os.path.abspath(root) == base_path:
            for d in dirs:
                if d.lower() not in ('sc', 'landscape', 'landscape rotate', 'edit', 'thumbnails', 'edit thumbnails'):
                    target_dirs.add(os.path.abspath(os.path.join(root, d)))

    # 3. Match root shortcuts to project folders using O(N * depth) walk
    print(f"Matching {len(cached)} root shortcuts to {len(target_dirs)} folders...")
    cached_map = {}
    target_dirs_set = {d.lower(): d for d in target_dirs}
    
    for c in cached:
        curr = c['target']
        while curr:
            curr_lower = curr.lower()
            if curr_lower in target_dirs_set:
                actual_dir = target_dirs_set[curr_lower]
                if actual_dir not in cached_map or c['date'] > cached_map[actual_dir]:
                    cached_map[actual_dir] = c['date']
                break
            parent = os.path.dirname(curr)
            if parent == curr: break # Root reached
            curr = parent

    # 4. Update each directory
    print("Finalizing updates...")
    updated_count = 0
    for directory in sorted(target_dirs):
        out_file = os.path.join(directory, 'scdate.txt')
        newest = datetime.min
        
        # Check project-local sc folder
        p_sc = os.path.join(directory, 'sc')
        if os.path.exists(p_sc):
            lnks = [os.path.join(p_sc, f) for f in os.listdir(p_sc) if f.lower().endswith('.lnk')]
            if lnks:
                latest_lnk = max(lnks, key=os.path.getmtime)
                newest = datetime.fromtimestamp(os.path.getmtime(latest_lnk), timezone.utc).replace(tzinfo=None)
        
        # Merge with root-level shortcut dates
        if directory in cached_map:
            if cached_map[directory] > newest:
                newest = cached_map[directory]
        
        if newest > datetime.min:
            write = True
            updated_count += 1
            if os.path.exists(out_file):
                try:
                    with open(out_file, 'r') as f:
                        content = f.read().strip()
                        d_date_str = content
                        if content.startswith('dummy:'):
                            d_date_str = content[6:].strip()
                        # ISO format: yyyy-MM-ddTHH:mm:ss.fffZ
                        # Python's fromisoformat might need a little help with the Z
                        if d_date_str.endswith('Z'):
                            d_date_str = d_date_str[:-1] + '+00:00'
                        d_date = datetime.fromisoformat(d_date_str).replace(tzinfo=None)
                        if newest <= d_date:
                            write = False
                except Exception:
                    pass
            
            if write:
                iso_date = newest.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
                if verbose:
                    print(f"  - {os.path.relpath(directory, base_path)}: {iso_date}")
                with open(out_file, 'w', encoding='utf-8') as f:
                    f.write(iso_date)
    print(f"Scanned {len(target_dirs)} directories, updated {updated_count} scdate.txt files.")

def update_sc_data(verbose=False):
    print("\nUpdating scdata.txt files...")
    base_path = os.getcwd()
    skip_dirs = {'.git', '__pycache__', 'thumbnails', 'edit thumbnails', '$recycle.bin', 'system volume information', 'originals'}
    # Recursive sc folders
    for root, dirs, files in os.walk(base_path):
        dirs[:] = [d for d in dirs if d.lower() not in skip_dirs]
        if 'sc' in dirs:
            sc_path = os.path.join(root, 'sc')
            links = [f for f in os.listdir(sc_path) if f.lower().endswith('.lnk')]
            if links:
                out = os.path.join(root, 'scdata.txt')
                if verbose:
                    print(f"  - {os.path.relpath(root, base_path)}: {len(links)} shortcuts")
                with open(out, 'w', encoding='utf-8') as f:
                    for l in sorted(links):
                        f.write(l + '\n')

    # Top-level ".\sc" (grouped target output)
    root_sc = os.path.join(base_path, 'sc')
    out = os.path.join(base_path, 'rootdata.txt')
    if os.path.exists(root_sc):
        groups = {}
        lnk_paths = [os.path.join(root_sc, f) for f in os.listdir(root_sc) if f.lower().endswith('.lnk')]
        targets_map = get_shortcut_targets_bulk(lnk_paths)
        update_root_misc_shortcuts(base_path, targets_map, verbose=verbose)
        if lnk_paths:
            for f_name in sorted(os.listdir(root_sc)):
                p = os.path.join(root_sc, f_name)
                t = targets_map.get(p)
                if t:
                    folder_path = os.path.dirname(t)
                    folder = os.path.basename(folder_path)
                    tag = '[ROOT]'
                    sub_sc = os.path.join(folder_path, 'sc')
                    if os.path.exists(os.path.join(sub_sc, f_name)):
                        tag = '[BOTH]'
                    if folder not in groups:
                        groups[folder] = []
                    groups[folder].append(f"{f_name} {tag}")
        
        if groups:
            if verbose:
                print(f"  - root: {sum(len(g) for g in groups.values())} shortcuts in {len(groups)} groups")
            with open(out, 'w', encoding='utf-8') as f:
                for folder in sorted(groups.keys()):
                    f.write(f'"{folder}"\n')
                    for entry in sorted(groups[folder]):
                        f.write(entry + '\n')
                    f.write('\n')
        elif os.path.exists(out):
            os.remove(out)
    elif os.path.exists(out):
        os.remove(out)

def generate_sc_new(verbose=False):
    print("\nGenerating scnew.txt...")
    base_path = os.getcwd()
    root_sc = os.path.join(base_path, 'sc')
    root_links_data = []
    if os.path.exists(root_sc):
        lnk_paths = [os.path.join(root_sc, f) for f in os.listdir(root_sc) if f.lower().endswith('.lnk')]
        if lnk_paths:
            targets_map = get_shortcut_targets_bulk(lnk_paths)
            for p, t in targets_map.items():
                root_links_data.append({'path': p, 'target': t, 'mtime': os.path.getmtime(p)})

    for d in os.listdir(base_path):
        dp = os.path.join(base_path, d)
        if os.path.isdir(dp) and d != 'sc':
            proj_sc = os.path.join(dp, 'sc')
            scnew_file = os.path.join(dp, 'scnew.txt')
            
            if not os.path.exists(proj_sc):
                if os.path.exists(scnew_file): os.remove(scnew_file)
                continue
            
            proj_links = [os.path.join(proj_sc, f) for f in os.listdir(proj_sc) if f.lower().endswith('.lnk')]
            if not proj_links:
                if os.path.exists(scnew_file): os.remove(scnew_file)
                continue
            
            matching_root_links = [l for l in root_links_data if l['target'].lower().startswith(dp.lower() + os.sep) or l['target'].lower() == dp.lower()]
            
            new_links = []
            if not matching_root_links:
                new_links = sorted(proj_links, key=os.path.getmtime)
            else:
                cutoff = max(l['mtime'] for l in matching_root_links)
                new_links = [l for l in proj_links if os.path.getmtime(l) > cutoff]
                new_links.sort(key=os.path.getmtime)
            
            if new_links:
                if verbose:
                    print(f"  - {d}: {len(new_links)} new shortcuts")
                with open(scnew_file, 'w', encoding='utf-8') as f:
                    for l in new_links:
                        f.write(os.path.basename(l) + '\n')
            else:
                if os.path.exists(scnew_file): os.remove(scnew_file)

def update_selections(verbose=False):
    print("\nUpdating selections.txt files...")
    base_path = os.getcwd()
    special_folders = {'sc', 'landscape', 'landscape rotate', 'edit', 'thumbnails', 'edit thumbnails', 'originals'}
    
    for d in os.listdir(base_path):
        dp = os.path.join(base_path, d)
        if os.path.isdir(dp) and d.lower() not in special_folders:
            if verbose:
                print(f"  - {d}")
            out = os.path.join(dp, 'selections.txt')
            with open(out, 'w', encoding='utf-8') as f:
                for sub in ["sc", "Landscape", "Landscape Rotate", "Edit"]:
                    f.write(f"# {sub}\n")
                    sub_path = os.path.join(dp, sub)
                    if os.path.exists(sub_path):
                        items = sorted(os.listdir(sub_path))
                        for item in items:
                            if os.path.isfile(os.path.join(sub_path, item)):
                                f.write(item + '\n')
                    f.write("\n")

def update_shortcut_database(verbose=False):
    print("\nUpdating Shortcut Database...")
    base_path = os.path.abspath(os.getcwd())
    db_file = "shortcut_db.txt"
    database = []
    skip_dirs = {'.git', '__pycache__', 'thumbnails', 'edit thumbnails', '$recycle.bin', 'system volume information', 'originals'}
    if os.path.exists(db_file):
        with open(db_file, 'r') as f:
            current_entry = {}
            for line in f:
                line = line.strip()
                if line.startswith('Folder path: '): current_entry['FolderPath'] = line[13:]
                elif line.startswith('Shortcut: '): current_entry['ShortcutName'] = line[10:]
                elif line.startswith('Shortcut Video Path: '): current_entry['VideoPath'] = line[21:]
                elif line.startswith('Shortcut md5: '): current_entry['MD5'] = line[14:]
                elif line == '---':
                    if 'FolderPath' in current_entry: database.append(current_entry)
                    current_entry = {}

    new_database = []
    # Collect all shortcuts first
    all_lnk_paths = []
    for root, dirs, files in os.walk(base_path):
        dirs[:] = [d for d in dirs if d.lower() not in skip_dirs]
        for file in files:
            if file.lower().endswith('.lnk'):
                all_lnk_paths.append(os.path.join(root, file))
    
    # Bulk resolve
    targets_map = get_shortcut_targets_bulk(all_lnk_paths)

    for lnk_path in all_lnk_paths:
        file = os.path.basename(lnk_path)
        root = os.path.dirname(lnk_path)
        target = targets_map.get(lnk_path)
        if not target: continue
        
        if not target.lower().endswith(VIDEO_EXTENSIONS): continue
        
        existing = next((e for e in database if e['FolderPath'] == root and e['ShortcutName'] == file), None)
        if existing:
            if os.path.exists(target):
                existing['VideoPath'] = target
                existing['MD5'] = get_md5(target)
            new_database.append(existing)
        else:
            if os.path.exists(target):
                md5 = get_md5(target)
                new_database.append({
                    'FolderPath': root,
                    'ShortcutName': file,
                    'VideoPath': target,
                    'MD5': md5
                })
                if verbose:
                    print(f"  - Added: {file}")
                else:
                    print(f"Added: {file}")

    with open(db_file, 'w', encoding='utf-8') as f:
        for entry in new_database:
            f.write(f"Folder path: {entry['FolderPath']}\n")
            f.write(f"Shortcut: {entry['ShortcutName']}\n")
            f.write(f"Shortcut Video Path: {entry['VideoPath']}\n")
            f.write(f"Shortcut md5: {entry['MD5']}\n")
            f.write("---\n")
    print(f"Database updated. Total entries: {len(new_database)}")

def scan_broken_shortcuts_from_db():
    print("\nScanning for broken shortcuts...")
    db_file = "shortcut_db.txt"
    if not os.path.exists(db_file):
        print("Shortcut database not found. Please run 'Update shortcut Database' first.")
        return

    database = []
    with open(db_file, 'r') as f:
        current_entry = {}
        for line in f:
            line = line.strip()
            if line.startswith('Folder path: '): current_entry['FolderPath'] = line[13:]
            elif line.startswith('Shortcut: '): current_entry['ShortcutName'] = line[10:]
            elif line.startswith('Shortcut Video Path: '): current_entry['VideoPath'] = line[21:]
            elif line.startswith('Shortcut md5: '): current_entry['MD5'] = line[14:]
            elif line == '---':
                if 'FolderPath' in current_entry: database.append(current_entry)
                current_entry = {}

    for entry in database:
        lnk_path = os.path.join(entry['FolderPath'], entry['ShortcutName'])
        if not os.path.exists(lnk_path): continue
        
        target = get_shortcut_target(lnk_path)
        if target and os.path.exists(target): continue
        
        print(f"\nBroken Shortcut found: {entry['ShortcutName']} in {entry['FolderPath']}")
        print(f"Original Target: {entry['VideoPath']}")
        
        original_dir = os.path.dirname(entry['VideoPath'])
        if os.path.exists(original_dir):
            print(f"Searching for matching file in: {original_dir}")
            found_match = None
            for f in os.listdir(original_dir):
                f_path = os.path.join(original_dir, f)
                if os.path.isfile(f_path) and f.lower().endswith(VIDEO_EXTENSIONS):
                    if get_md5(f_path) == entry['MD5']:
                        found_match = f_path
                        break
            
            if found_match:
                print(f"Match found! New file name: {os.path.basename(found_match)}")
                if input("Repair shortcut? (y/n): ").lower() == 'y':
                    if update_shortcut(lnk_path, found_match):
                        print("Shortcut repaired.")
            else:
                print("No matching file found by MD5 in the original directory.")
        else:
            print(f"Original target directory no longer exists: {original_dir}")

def run_shortcut_manager_menu(verbose=False):
    while True:
        print("\n--- Shortcut Manager ---")
        print("1. Update shortcut Database")
        print("2. Scan for broken Shortcuts")
        print("3. Back")
        choice = input("\nSelect an option: ")
        if choice == '1': update_shortcut_database(verbose=verbose)
        elif choice == '2': scan_broken_shortcuts_from_db()
        elif choice == '3': break
        else: print("Invalid choice.")

def run_broken_shortcuts_scan(generate_report=False):
    base_path = os.getcwd()
    projects = get_projects(base_path)
    broken_shortcuts = [] # list of (shortcut_path, current_target)
    skip_dirs = {'.git', '__pycache__', 'thumbnails', 'edit thumbnails', '$recycle.bin', 'system volume information', 'originals'}

    def scan_dir_for_shortcuts(directory, shortcuts_to_check):
        if not os.path.isdir(directory): return
        for f in os.listdir(directory):
            p = os.path.join(directory, f)
            if is_shortcut(p):
                shortcuts_to_check.append(p)

    print("\nScanning for broken shortcuts...")
    all_lnk_to_check = []
    # Scan current directory
    scan_dir_for_shortcuts(base_path, all_lnk_to_check)
    # Scan root 'sc' folder if it exists
    scan_dir_for_shortcuts(os.path.join(base_path, 'sc'), all_lnk_to_check)
    # Scan each project's 'sc' folder
    for project in projects:
        scan_dir_for_shortcuts(os.path.join(base_path, project, 'sc'), all_lnk_to_check)
    
    if all_lnk_to_check:
        targets_map = get_shortcut_targets_bulk(all_lnk_to_check)
        for p in all_lnk_to_check:
            target = targets_map.get(p)
            if not target or not os.path.exists(target):
                broken_shortcuts.append((p, target))

    if not broken_shortcuts:
        print("No broken shortcuts found.")
        return

    print(f"Found {len(broken_shortcuts)} broken shortcuts.")

    print("\nBroken Shortcuts:")
    for p, t in broken_shortcuts:
        print(f"  - {os.path.relpath(p, base_path)} -> {t}")

    if generate_report:
        with open('broken_shortcuts_report.txt', 'w', encoding='utf-8') as f:
            f.write(f"BROKEN SHORTCUTS REPORT - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("========================================================\n\n")
            for p, t in broken_shortcuts:
                f.write(f"{p} -> {t}\n")
        print("broken_shortcuts_report.txt generated.")

    if input("\nWould you like to search for correct paths? (y/n): ").lower() == 'y':
        print("\nSearching for correct paths...")
        # Collect all videos from all projects once
        all_videos = []
        for project in projects:
            all_videos.extend(find_videos(os.path.join(base_path, project)))
        
        # filename -> [full_paths]
        video_map = {}
        for v in all_videos:
            name = os.path.basename(v)
            if name not in video_map: video_map[name] = []
            video_map[name].append(v)

        for p, t in broken_shortcuts:
            # Shortcut name (without extension if it's .lnk)
            sc_name = os.path.basename(p)
            if sc_name.lower().endswith('.lnk'):
                # Try both the name with and without .lnk if needed, 
                # but usually shortcut matches filename or filename.lnk
                # We'll use the base filename part for matching.
                search_name = os.path.splitext(sc_name)[0]
            else:
                search_name = sc_name
            
            matches = []
            # Exact match for search_name in video names
            for v_name, paths in video_map.items():
                if v_name == search_name or os.path.splitext(v_name)[0] == search_name:
                    matches.extend(paths)
            
            if matches:
                print(f"\nBroken shortcut: {os.path.relpath(p, base_path)}")
                print(f"Current invalid target: {t}")
                print("Found matching video(s):")
                for i, match in enumerate(matches):
                    print(f"  {i+1}. {os.path.relpath(match, base_path)}")
                
                choice = input("Enter number to fix shortcut, or 's' to skip: ")
                if choice.isdigit() and 1 <= int(choice) <= len(matches):
                    new_target = matches[int(choice)-1]
                    if update_shortcut(p, new_target):
                        print("Shortcut updated successfully.")
                else:
                    print("Skipped.")
            else:
                print(f"\nNo matches found for broken shortcut: {os.path.relpath(p, base_path)}")

def run_empty_video_scan(generate_report=False):
    base_path = os.getcwd()
    projects = get_projects(base_path)
    empty_videos = []
    
    print("\nScanning for empty videos (< 2KB)...")
    for project in projects:
        project_path = os.path.join(base_path, project)
        videos = find_videos(project_path)
        for video in videos:
            try:
                if os.path.getsize(video) < 2048:
                    empty_videos.append(video)
            except OSError:
                continue

    if not empty_videos:
        print("No empty videos found.")
        return

    print(f"Found {len(empty_videos)} empty videos.")
    
    print("\nEmpty Videos:")
    for v in empty_videos:
        print(f"  - {os.path.relpath(v, base_path)}")
            
    if generate_report:
        with open('empty_videos_report.txt', 'w', encoding='utf-8') as f:
            f.write(f"EMPTY VIDEOS REPORT - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("========================================================\n\n")
            for v in empty_videos:
                f.write(f"{v}\n")
        print("empty_videos_report.txt generated.")

    if input("\nWould you like to delete these empty videos? (y/n): ").lower() == 'y':
        deleted_count = 0
        for v in empty_videos:
            try:
                os.remove(v)
                deleted_count += 1
            except Exception as e:
                print(f"Error deleting {v}: {e}")
        print(f"Deleted {deleted_count} empty videos.")


def run_normal_scan(deep_scan=False, generate_report=False):
    base_path = os.getcwd()
    projects = get_projects(base_path)
    
    total_projects = len(projects)
    results = {
        'total_videos': 0, 'total_unique_names': 0,
        'total_found_main_slots': 0, 'total_found_edit_slots': 0,
        'total_unique_main_files': 0, 'total_unique_edit_files': 0,
        'total_wrong_dimensions': 0,
    }
    
    project_reports = []
    generation_queue = {}
    obsolete_images = [] # List of (absolute_path)
    
    print()
    for idx, project in enumerate(projects):
        project_path = os.path.join(base_path, project)
        percent = (idx / total_projects) * 100
        
        thumb_dir = os.path.join(project_path, 'Thumbnails')
        edit_dir = os.path.join(project_path, 'Edit Thumbnails')
        thumb_files = set(os.listdir(thumb_dir)) if os.path.isdir(thumb_dir) else set()
        edit_files = set(os.listdir(edit_dir)) if os.path.isdir(edit_dir) else set()
        
        videos = find_videos(project_path)
        project_videos_report = []
        proj_all_names, proj_found_main_names, proj_found_edit_slots_set = set(), set(), set()
        proj_found_main_slots, proj_found_edit_slots = 0, 0
        
        used_thumb_files = set()
        used_edit_files = set()
        
        if not videos:
            print(f"[{percent:6.2f}%] Scanning project: {project} - No videos found.        ", end='\r')
            # Even if no videos, we should check if there are images in Thumbnails/Edit Thumbnails
            # because they'd all be obsolete.
            for f in thumb_files:
                if any(f.lower().endswith(ext.lower()) for ext in IMAGE_EXTENSIONS):
                    obsolete_images.append(os.path.join(thumb_dir, f))
            for f in edit_files:
                if any(f.lower().endswith(ext.lower()) for ext in IMAGE_EXTENSIONS):
                    obsolete_images.append(os.path.join(edit_dir, f))
            continue
            
        for v_idx, video in enumerate(videos):
            video_rel_path = os.path.relpath(video, project_path)
            video_name = os.path.splitext(os.path.basename(video))[0]
            proj_all_names.add(video_name)
            
            overall_percent = ((idx + (v_idx / len(videos))) / total_projects) * 100
            print(f"[{overall_percent:6.2f}%] Project: {project} | Scanning: {video_rel_path}                ", end='\r')
            
            main_file, edit_indices, edit_files_found = check_thumbnails_optimized(video, project_path, thumb_files, edit_files)
            
            # Dimension checking for Deep Scan
            wrong_dim_edits = []
            needs_main_fix = False
            
            if deep_scan:
                nb_frames, fps, v_width, v_height, v_ar = get_video_info(video)
                if v_width > 0 and v_height > 0:
                    diffs = {
                        'portrait': (abs(v_ar - 0.5625), 284, 504),
                        'square': (abs(v_ar - 1.0), 504, 504),
                        'landscape': (abs(v_ar - 1.777), 896, 504)
                    }
                    group_name, (_, target_w, target_h) = min(diffs.items(), key=lambda x: x[1][0])

                    if v_ar >= (target_w / target_h):
                        rw, rh = target_w, int(target_w / v_ar)
                    else:
                        rh, rw = target_h, int(target_h * v_ar)

                    w_diff_tol = abs(rw - target_w) / target_w
                    h_diff_tol = abs(rh - target_h) / target_h
                    needs_pad = w_diff_tol > 0.10 or h_diff_tol > 0.10

                    def is_valid_dim(w, h):
                        if w == target_w and h == target_h: return True
                        if not needs_pad:
                            if abs(w - rw) <= 2 and abs(h - rh) <= 2:
                                return True
                        return False

                    if main_file:
                        main_path = os.path.join(thumb_dir, main_file)
                        w, h = get_image_dimensions(main_path)
                        if not is_valid_dim(w, h):
                            needs_main_fix = True
                            results['total_wrong_dimensions'] += 1

                    for i, f in zip(edit_indices, edit_files_found):
                        edit_path = os.path.join(edit_dir, f)
                        w, h = get_image_dimensions(edit_path)
                        if not is_valid_dim(w, h):
                            wrong_dim_edits.append(i)
                            results['total_wrong_dimensions'] += 1

            if main_file:
                proj_found_main_slots += 1
                proj_found_main_names.add(video_name)
                used_thumb_files.add(main_file)
            proj_found_edit_slots += len(edit_indices)
            for i in edit_indices: proj_found_edit_slots_set.add((video_name, i))
            for f in edit_files_found: used_edit_files.add(f)
            
            missing_edits = [i for i in range(1, 11) if i not in edit_indices]
            needs_main = main_file is None or needs_main_fix
            missing_edits.extend(wrong_dim_edits)
            missing_edits = sorted(list(set(missing_edits)))
            
            if needs_main or missing_edits:
                if video not in generation_queue:
                    generation_queue[video] = [project_path, needs_main, missing_edits]
                else:
                    if needs_main: generation_queue[video][1] = True
                    # Combine missing edits
                    existing_missing = set(generation_queue[video][2])
                    existing_missing.update(missing_edits)
                    generation_queue[video][2] = sorted(list(existing_missing))
            
            project_videos_report.append({
                'video': video_rel_path, 'main_thumbnail': main_file is not None,
                'edit_thumbnails_count': len(edit_indices)
            })
            
        results['total_videos'] += len(videos)
        results['total_unique_names'] += len(proj_all_names)
        results['total_found_main_slots'] += proj_found_main_slots
        results['total_found_edit_slots'] += proj_found_edit_slots
        results['total_unique_main_files'] += len(proj_found_main_names)
        results['total_unique_edit_files'] += len(proj_found_edit_slots_set)
        
        # Identify obsolete images in this project
        for f in thumb_files:
            if f not in used_thumb_files and any(f.lower().endswith(ext.lower()) for ext in IMAGE_EXTENSIONS):
                obsolete_images.append(os.path.join(thumb_dir, f))
        for f in edit_files:
            if f not in used_edit_files and any(f.lower().endswith(ext.lower()) for ext in IMAGE_EXTENSIONS):
                obsolete_images.append(os.path.join(edit_dir, f))

        project_reports.append({
            'name': project, 'videos': project_videos_report,
            'main_thumbs_found': proj_found_main_slots, 'main_thumbs_expected': len(videos),
            'edit_thumbs_found': proj_found_edit_slots, 'edit_thumbs_expected': len(videos) * 10,
            'missing': (len(proj_all_names) - len(proj_found_main_names)) + (len(proj_all_names) * 10 - len(proj_found_edit_slots_set))
        })

    print(f"\nScan complete.\n")
    total_missing = (results['total_unique_names'] - results['total_unique_main_files']) + \
                    (results['total_unique_names'] * 10 - results['total_unique_edit_files'])
    summary = (
        f"Total number of projects scanned: {total_projects}\n"
        f"Total number of videos found: {results['total_videos']} ({results['total_unique_names']} unique names)\n"
        f"Main Thumbs: Found {results['total_found_main_slots']}/{results['total_videos']} ({results['total_unique_main_files']} unique files)\n"
        f"Edit Thumbs: Found {results['total_found_edit_slots']}/{results['total_videos'] * 10} ({results['total_unique_edit_files']} unique files)\n"
    )
    if deep_scan:
        summary += f"Images with wrong dimensions: {results['total_wrong_dimensions']}\n"
    summary += (
        f"Total number of missing images: {total_missing}\n"
        f"Total number of obsolete images: {len(obsolete_images)}\n"
    )
    print(summary)
    
    if total_missing > 0:
        print("\nFiles with Missing Images:")
        for pr in project_reports:
            missing_in_project = [vr for vr in pr['videos'] if not vr['main_thumbnail'] or vr['edit_thumbnails_count'] < 10]
            if missing_in_project:
                print(f"\n[{pr['name']}]")
                for vr in missing_in_project:
                    status = []
                    if not vr['main_thumbnail']: status.append("Main")
                    if vr['edit_thumbnails_count'] < 10: status.append(f"Edits ({vr['edit_thumbnails_count']}/10)")
                    print(f"  - {vr['video']} | Missing: {', '.join(status)}")
        print()
    
    if obsolete_images:
        if input(f"Would you like to delete {len(obsolete_images)} obsolete images? (y/n): ").lower() == 'y':
            for img_path in obsolete_images:
                try:
                    os.remove(img_path)
                except Exception as e:
                    print(f"Error deleting {img_path}: {e}")
            print(f"Deleted {len(obsolete_images)} obsolete images.")

    if generation_queue:
        gen_main_count = sum(1 for v in generation_queue.values() if v[1])
        gen_edit_count = sum(1 for v in generation_queue.values() if v[2])
        print(f"Eligible for thumbnail generation:\n- {gen_main_count} Main, {gen_edit_count} Edit videos")
        
        if input("Would you like to generate missing thumbnails? (y/n): ").lower() == 'y':
            try: subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True)
            except Exception:
                print("Error: FFmpeg not found.")
                return

            tasks = [(path, data[0], data[1], data[2]) for path, data in generation_queue.items()]
            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(generate_video_thumbnails, task) for task in tasks]
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    video_path, success = future.result()
                    v_rel = os.path.relpath(video_path, base_path)
                    print(f"[{i+1}/{len(tasks)}] Processed: {v_rel}            ", end='\r')
            print(f"\nGeneration complete.\n")

    if generate_report:
        with open('report.txt', 'w') as f:
            f.write(f"THUMBNAIL SCAN REPORT - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("========================================================\n\n")
            f.write(summary)
            for pr in project_reports:
                f.write(f"\nProject: {pr['name']} | Missing: {pr['missing']}\n")
                for vr in pr['videos']:
                    f.write(f"  - {vr['video']} | Main: {'OK' if vr['main_thumbnail'] else 'MISSING'} | Edits: {vr['edit_thumbnails_count']}/10\n")
        print("report.txt generated.")

def main():
    while True:
        print("\033[1;33m") # Bold Yellow
        print("======================================")
        print("   SC Utilities")
        print("======================================")
        print("\033[0m") # Reset

        print() # Padding
        print("1. Normal Scan (Scan projects for missing and obsolete thumbnails)")
        print("2. Deep Scan (Scan projects for missing and obsolete or incorrect thumbnails)")
        print("-" * 38)
        print("3. Update scdate.txt (newest shortcut date)")
        print("4. Update scdata.txt (shortcut data)")
        print("5. Generate scnew.txt (for Load SC New)")
        print("6. Update selections.txt")
        print("7. Perform ALL updates (3-7)")
        print("-" * 38)
        print("8. Shortcut Manager")
        print("-" * 38)
        print("9. Scan for empty videos")
        print("10. Scan for broken shortcuts")
        print("-" * 38)
        print("11. Exit")
        
        user_input = input("\nEnter choice(s) (e.g., 1, 4 or 1r): ")
        if not user_input.strip(): continue
        
        choices = re.split(r'[ ,]+', user_input.strip())
        
        for choice in choices:
            choice = choice.strip()
            if not choice: continue
            
            generate_report = False
            clean_choice = choice
            if choice.lower().endswith('r'):
                generate_report = True
                clean_choice = choice[:-1]
            
            if clean_choice == '1': run_normal_scan(deep_scan=False, generate_report=generate_report)
            elif clean_choice == '2': run_normal_scan(deep_scan=True, generate_report=generate_report)
            elif clean_choice == '3': update_sc_date(verbose=generate_report)
            elif clean_choice == '4': update_sc_data(verbose=generate_report)
            elif clean_choice == '5': generate_sc_new(verbose=generate_report)
            elif clean_choice == '6': update_selections(verbose=generate_report)
            elif clean_choice == '7':
                update_sc_date(verbose=generate_report)
                update_sc_data(verbose=generate_report)
                generate_sc_new(verbose=generate_report)
                update_selections(verbose=generate_report)
            elif clean_choice == '8': run_shortcut_manager_menu(verbose=generate_report)
            elif clean_choice == '9': run_empty_video_scan(generate_report=generate_report)
            elif clean_choice == '10': run_broken_shortcuts_scan(generate_report=generate_report)
            elif clean_choice == '11': return
            else: print(f"Invalid choice: {choice}")

if __name__ == "__main__":
    main()
