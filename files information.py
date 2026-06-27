import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import win32security
import win32con

def get_file_info(filepath):
    try:
        stats = os.stat(filepath)
        creation_time = datetime.datetime.fromtimestamp(stats.st_ctime)
        last_modified_time = datetime.datetime.fromtimestamp(stats.st_mtime)
        
        # Get owner information
        try:
            sd = win32security.GetFileSecurity(filepath, win32security.OWNER_SECURITY_INFORMATION)
            owner_sid = sd.GetSecurityDescriptorOwner()
            name, domain, _ = win32security.LookupAccountSid(None, owner_sid)
            owner = f"{domain}\\{name}"
        except Exception:
            owner = 'Unavailable'
        
        # Get creator/modifier information
        try:
            sd = win32security.GetFileSecurity(filepath, win32security.DACL_SECURITY_INFORMATION)
            aces = sd.GetSecurityDescriptorDacl()
            
            creator = modifier = None
            for ace_index in range(aces.GetAceCount()):
                ace = aces.GetAce(ace_index)
                if ace[0][1] & win32con.WRITE_OWNER:
                    if not modifier:
                        modifier = win32security.LookupAccountSid(None, ace[2])[0]
                    if not creator:
                        creator = win32security.LookupAccountSid(None, ace[2])[0]
            
            created_by = creator if creator else owner
            modified_by = modifier if modifier else owner
            
            created_info = f"{created_by} ({creation_time.strftime('%Y-%m-%d %H:%M:%S')})"
            modified_info = f"{modified_by} ({last_modified_time.strftime('%Y-%m-%d %H:%M:%S')})"
            
        except Exception:
            created_info = f"{owner} ({creation_time.strftime('%Y-%m-%d %H:%M:%S')})"
            modified_info = f"{owner} ({last_modified_time.strftime('%Y-%m-%d %H:%M:%S')})"

        return {
            'File Name': os.path.basename(filepath),
            'Full Path': filepath,
            'Owner': owner,
            'Created': created_info,
            'Last Modified': modified_info
        }
        
    except Exception as e:
        return {
            'File Name': os.path.basename(filepath),
            'Full Path': filepath,
            'Owner': f'Error: {str(e)}',
            'Created': 'Error',
            'Last Modified': 'Error'
        }

def main():
    try:
        root = tk.Tk()
        root.withdraw()
        
        # Ask for folder
        folder_path = filedialog.askdirectory(title='Select Folder to Scan')
        if not folder_path:
            messagebox.showinfo("Cancelled", "No folder selected. Operation cancelled.")
            return
        
        # Collect file information
        files_data = []
        total_files = 0
        for root_dir, _, files in os.walk(folder_path):
            for filename in files:
                filepath = os.path.join(root_dir, filename)
                files_data.append(get_file_info(filepath))
                total_files += 1
                if total_files % 100 == 0:  # Update progress every 100 files
                    print(f"Processed {total_files} files...")
        
        if not files_data:
            messagebox.showwarning("Warning", "No files found in the selected folder.")
            return
        
        # Create DataFrame
        df = pd.DataFrame(files_data)
        
        # Configure display for full output
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', None)
        
        print("\nFile Information Summary:")
        print(df)
        
        # Ask for save location
        save_path = filedialog.asksaveasfilename(
            title='Save As Excel File',
            defaultextension='.xlsx',
            filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')],
            initialfile='File_Information.xlsx'
        )
        
        if save_path:
            # Save to Excel with auto-adjusted columns
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='File Info')
                
                # Auto-adjust column widths
                worksheet = writer.sheets['File Info']
                for i, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, max_len)
            
            messagebox.showinfo("Success", f"Successfully processed {total_files} files.\nSaved to:\n{save_path}")
        
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred:\n{str(e)}")

if __name__ == "__main__":
    main()