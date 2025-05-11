import pandas as pd
import os
import re
from pathlib import Path

def normalize_path(path):
    """
    Normalizes a path string to handle different path formats and separators.
    
    Args:
        path (str): The path to normalize
    Returns:
        str: Normalized absolute path
    """
    # Remove quotes if present
    path = path.strip('"\'')
    # Convert to Path object and resolve to absolute path
    return str(Path(path).resolve())

def read_file(file_path):
    """
    Reads a file based on its extension (csv or xlsx)
    
    Args:
        file_path (str): Path to the file to read
    Returns:
        pd.DataFrame: DataFrame containing the file data
    """
    file_path = normalize_path(file_path)
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.csv':
        return pd.read_csv(file_path)
    elif file_extension == '.xlsx':
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}. Please use .csv or .xlsx files.")

def save_file(df, file_path):
    """
    Saves a DataFrame to a file based on its extension (csv or xlsx)
    
    Args:
        df (pd.DataFrame): DataFrame to save
        file_path (str): Path where to save the file
    """
    file_path = normalize_path(file_path)
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.csv':
        df.to_csv(file_path, index=False)
    elif file_extension == '.xlsx':
        df.to_excel(file_path, index=False)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}. Please use .csv or .xlsx files.")

def remove_column_numbering(df):
    """
    Removes trailing decimal numbers from column names while preserving the data.
    Example: Converts columns like:
    'Custom field (Using Legal Entity (Application)).1', 
    'Custom field (Using Legal Entity (Application)).2'
    to:
    'Custom field (Using Legal Entity (Application))',
    'Custom field (Using Legal Entity (Application))'
    
    Args:
        df (pd.DataFrame): DataFrame with numbered columns
    Returns:
        pd.DataFrame: DataFrame with cleaned column names
    """
    # Create a mapping of old column names to new ones (without the trailing numbers)
    column_mapping = {}
    for col in df.columns:
        # Remove trailing .number pattern
        new_col = re.sub(r'\.\d+$', '', col)
        column_mapping[col] = new_col
    
    # Rename the columns while preserving the data
    df.columns = [column_mapping[col] for col in df.columns]
    return df

def standardize_data(base_file, new_file):
    """
    Standardizes the structure of a new data file to match the base file.
    Handles repeated column headers (instances) and preserves their data.
    Only adds number suffixes to columns with multiple instances.
    Adds missing columns with empty values.
    
    Args:
        base_file (str): Path to the base file (csv or xlsx)
        new_file (str): Path to the new file to standardize (csv or xlsx)
    """
    try:
        # Read the base file to get the reference structure
        base_df = read_file(base_file)
        base_columns = base_df.columns.tolist()
        
        # Count instances of each column in base file
        base_column_counts = {}
        for col in base_columns:
            base_column_counts[col] = base_column_counts.get(col, 0) + 1
        
        # Read the new file
        new_df = read_file(new_file)
        new_columns = new_df.columns.tolist()
        
        # Count instances of each column in new file
        new_column_counts = {}
        for col in new_columns:
            new_column_counts[col] = new_column_counts.get(col, 0) + 1
        
        # Check if any column has fewer instances in new file
        missing_instances = []
        for col, count in base_column_counts.items():
            if col not in new_column_counts or new_column_counts[col] < count:
                missing_instances.append((col, count, new_column_counts.get(col, 0)))
        
        if missing_instances:
            print("WARNING: The following columns have insufficient instances in the new file:")
            for col, required, actual in missing_instances:
                print(f"- {col}: Required {required} instances, found {actual} instances")
            print("\nEmpty columns will be added for missing instances.")
        
        # Create standardized DataFrame
        standardized_df = pd.DataFrame()
        
        # Process each unique column (removing duplicates from base_columns)
        unique_columns = []
        processed_columns = set()
        
        for col in base_columns:
            if col not in processed_columns:
                unique_columns.append(col)
                processed_columns.add(col)
        
        # Process each column and its instances
        for col in unique_columns:
            # Get all instances of this column from new file
            new_col_instances = [c for c in new_columns if c == col]
            required_count = base_column_counts[col]
            
            # If column has only one instance, don't add number suffix
            if required_count == 1:
                if new_col_instances:
                    standardized_df[col] = new_df[new_col_instances[0]]
                else:
                    # Add empty column if missing
                    standardized_df[col] = pd.Series(dtype='object')
            else:
                # Add each instance with number suffix
                for i in range(required_count):
                    if i < len(new_col_instances):
                        standardized_df[f"{col}_{i+1}"] = new_df[new_col_instances[i]]
                    else:
                        # Add empty column for missing instances
                        standardized_df[f"{col}_{i+1}"] = pd.Series(dtype='object')
        
        # Remove trailing numbers from column names before saving
        standardized_df = remove_column_numbering(standardized_df)
        
        try:
            # Save the standardized data back to the original file
            save_file(standardized_df, new_file)
            print(f"File {new_file} has been standardized successfully!")
            print("Column instances preserved as per original data.")
            print("Missing columns added with empty values.")
            print("Trailing decimal numbers removed from column names.")
        except PermissionError:
            print(f"\nERROR: Could not save the file '{new_file}'.")
            print("Please make sure that:")
            print("1. The file is not currently open in Excel")
            print("2. You have write permissions for the file")
            print("3. The file is not set to read-only")
            print("\nPlease close the file if it's open and try again.")
            return
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        if "Permission denied" in str(e):
            print("\nPlease make sure that:")
            print("1. The file is not currently open in Excel")
            print("2. You have write permissions for the file")
            print("3. The file is not set to read-only")
            print("\nPlease close the file if it's open and try again.")

def standardize_folder(base_file, folder_path):
    """
    Standardizes all files in a folder according to the base file structure.
    
    Args:
        base_file (str): Path to the base file (csv or xlsx)
        folder_path (str): Path to the folder containing files to standardize
    """
    try:
        # Verify base file exists
        if not os.path.exists(base_file):
            raise FileNotFoundError(f"Base file not found: {base_file}")
            
        # Verify folder exists
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"Folder not found: {folder_path}")
            
        # Get all files in the folder
        folder = Path(folder_path)
        files_to_process = []
        
        # Collect all csv and excel files
        for file in folder.glob('*'):
            if file.suffix.lower() in ['.csv', '.xlsx']:
                files_to_process.append(str(file))
        
        if not files_to_process:
            print(f"No CSV or Excel files found in {folder_path}")
            return
            
        print(f"Found {len(files_to_process)} files to process")
        
        # Process each file
        for file_path in files_to_process:
            print(f"\nProcessing: {os.path.basename(file_path)}")
            try:
                standardize_data(base_file, file_path)
            except Exception as e:
                print(f"Error processing {file_path}: {str(e)}")
                continue
                
        print("\nFolder processing completed!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    while True:
        try:
            # Get base file path
            base_file = input("\nEnter the path to the base file (csv or xlsx): ").strip()
            if not base_file:
                print("Please enter a valid file path.")
                continue
            
            # Normalize and validate base file path
            try:
                base_file = normalize_path(base_file)
            except Exception as e:
                print(f"\nError with base file path: {str(e)}")
                continue
            
            # Validate base file exists and has correct extension
            if not os.path.exists(base_file):
                print(f"\nError: Base file not found at: {base_file}")
                print("Please make sure:")
                print("1. The file exists at the specified location")
                print("2. You have entered the correct path")
                print("3. The file extension is .csv or .xlsx")
                continue
                
            if not base_file.lower().endswith(('.csv', '.xlsx')):
                print("\nError: Base file must be a .csv or .xlsx file")
                continue
            
            # Get folder path
            folder_path = input("\nEnter the path to the folder containing files to standardize: ").strip()
            if not folder_path:
                print("Please enter a valid folder path.")
                continue
            
            # Normalize and validate folder path
            try:
                folder_path = normalize_path(folder_path)
            except Exception as e:
                print(f"\nError with folder path: {str(e)}")
                continue
            
            # Validate folder exists
            if not os.path.exists(folder_path):
                print(f"\nError: Folder not found at: {folder_path}")
                print("Please make sure:")
                print("1. The folder exists at the specified location")
                print("2. You have entered the correct path")
                continue
                
            if not os.path.isdir(folder_path):
                print(f"\nError: The path is not a folder: {folder_path}")
                continue
            
            # If all validations pass, proceed with standardization
            print("\nStarting standardization process...")
            print(f"Base file: {base_file}")
            print(f"Folder to process: {folder_path}")
            
            standardize_folder(base_file, folder_path)
            break
            
        except Exception as e:
            print(f"\nAn error occurred: {str(e)}")
            retry = input("\nWould you like to try again? (y/n): ").strip().lower()
            if retry != 'y':
                break 