import os
import shutil

# Anchor paths relative to the project root (one level up from this script in the 'scripts' folder)
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(script_dir, ".."))

src_dir = os.path.join(project_root, "履歷表 6A31黃熹澄")
dst_dir = os.path.join(project_root, "flat_portfolio")

if not os.path.exists(dst_dir):
    os.makedirs(dst_dir)
    print(f"Created destination directory: {dst_dir}")

count = 0
for root, dirs, files in os.walk(src_dir):
    for file in files:
        # Ignore hidden/system files
        if file.startswith('.'):
            continue
            
        file_path = os.path.join(root, file)
        
        # Calculate relative path to src_dir
        rel_path = os.path.relpath(root, src_dir)
        
        if rel_path == ".":
            # File is in the root of src_dir
            new_name = file
        else:
            # File is in a subdirectory
            path_parts = rel_path.split(os.sep)
            parent_folder_name = path_parts[-1]
            
            # Extract code prefix (e.g. P4_045 or P5_002) if possible
            has_code = False
            if "_" in parent_folder_name:
                parts = parent_folder_name.split("_")
                if len(parts) >= 2 and parts[0].startswith(("P4", "P5", "P6")) and parts[1].isdigit():
                    code = f"{parts[0]}_{parts[1]}"
                    has_code = True
                    
            if has_code and file.startswith(code):
                # File already starts with the code, so it is already unique and clear
                new_name = file
            elif parent_folder_name == "附件":
                new_name = f"附件_{file}"
            elif parent_folder_name in ["四年級", "五年級", "六年級"]:
                new_name = f"{parent_folder_name}_{file}"
            else:
                new_name = f"{parent_folder_name}_{file}"
                
        # Safe-guard: truncate filename if it is too long (OS limit is 255, we use 150 for absolute safety)
        if len(new_name) > 150:
            name_part, ext_part = os.path.splitext(new_name)
            allowed_len = 150 - len(ext_part)
            new_name = name_part[:allowed_len] + ext_part
            
        dst_path = os.path.join(dst_dir, new_name)
        shutil.copy2(file_path, dst_path)
        count += 1

print(f"\nSuccess! Flattened and copied {count} files to '{dst_dir}'.")
