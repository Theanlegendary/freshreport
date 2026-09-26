import os
import zipfile
import time
import sys

def make_zip(source_dir, output_zip_path):
    print(f"Creating zip archive at: {output_zip_path}")
    print(f"Source directory: {source_dir}")
    
    start_time = time.time()
    total_files = 0
    total_bytes = 0
    
    # Pre-count files
    all_files = []
    abs_output = os.path.abspath(output_zip_path)
    base_folder = os.path.basename(os.path.normpath(source_dir))
    
    for root, dirs, files in os.walk(source_dir):
        # Exclude output zip if inside source_dir
        for f in files:
            full_path = os.path.join(root, f)
            if os.path.abspath(full_path) == abs_output:
                continue
            all_files.append(full_path)
            
    print(f"Found {len(all_files)} files to compress.")
    
    try:
        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=4, allowZip64=True) as zf:
            for idx, file_path in enumerate(all_files):
                rel_path = os.path.relpath(file_path, os.path.dirname(os.path.normpath(source_dir)))
                try:
                    zf.write(file_path, rel_path)
                    total_files += 1
                    total_bytes += os.path.getsize(file_path)
                except Exception as e:
                    print(f"Warning: Could not add file {idx}: {repr(e)}")
                    
                if (idx + 1) % 500 == 0 or idx == len(all_files) - 1:
                    elapsed = time.time() - start_time
                    mb_processed = total_bytes / (1024 * 1024)
                    print(f"Progress: [{idx + 1}/{len(all_files)}] files ({mb_processed:.1f} MB) in {elapsed:.1f}s")
                    sys.stdout.flush()
    except Exception as e:
        import traceback
        print("ERROR IN ZIP CREATION:")
        traceback.print_exc()
        raise

    zip_size_mb = os.path.getsize(output_zip_path) / (1024 * 1024)
    elapsed = time.time() - start_time
    print(f"\nSuccessfully created {output_zip_path}!")
    print(f"Total files: {total_files}")
    print(f"Original size: {total_bytes / (1024*1024):.1f} MB")
    print(f"Compressed zip size: {zip_size_mb:.1f} MB")
    print(f"Time taken: {elapsed:.1f} seconds")

if __name__ == '__main__':
    source_dir = r"c:\Users\DELL\Desktop\daily_push"
    output_zip = r"c:\Users\DELL\Desktop\daily_push.zip"
    make_zip(source_dir, output_zip)
