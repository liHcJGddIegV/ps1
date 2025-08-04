import pandas as pd
import glob
import os

def merge_excel_files(folder_path, output_file, sheet_name=0):
    """
    Merges all Excel (.xlsx) files in the specified folder into a single Excel file.
    
    Args:
    - folder_path (str): The directory containing the Excel files.
    - output_file (str): Path for the merged output file.
    - sheet_name (int or str, optional): The sheet to merge from each file. Default is the first sheet (0).
    
    Returns:
    - None
    """
    all_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    
    if not all_files:
        print("No Excel files found in the folder. Exiting.")
        return

    df_list = []
    
    for file in all_files:
        try:
            df = pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")  # Specify engine explicitly
            df['Source_File'] = os.path.basename(file)  # Track file origin
            df_list.append(df)
            print(f"Processed: {file}")
        except Exception as e:
            print(f"Error reading {file}: {e}")

    if not df_list:
        print("No valid data found to merge. Exiting.")
        return
    
    merged_df = pd.concat(df_list, ignore_index=True)

    try:
        merged_df.to_excel(output_file, index=False, engine="openpyxl")
        print(f"Merging complete! Output saved to '{output_file}'")
    except Exception as e:
        print(f"Failed to save merged file: {e}")

if __name__ == "__main__":
    # Update these paths as needed
    folder_path = r"C:\Users\YGonzalez\OneDrive - Invenergy LLC\Documents\ProjectTeams"
    output_file = os.path.join(folder_path, "merged_output.xlsx")  # Save in the same directory
    
    merge_excel_files(folder_path, output_file)
