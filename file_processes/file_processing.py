import pandas as pd
import os
import re
class FileProcessor:
    def __init__(self, upload_folder):
        self.upload_folder = upload_folder
        
    #function to read the excel file based on the file extension
    def _read_excel(self, file_path, filename):
        try:
            extension = os.path.splitext(filename)[1].lower()
        
        # Choose the correct engine
            if extension == '.xls':
                engine = 'xlrd'
            elif extension == '.xlsx':
                engine = 'openpyxl'
            else:
                raise ValueError("Unsupported file format. Only .xls and .xlsx are allowed.")
        
            df = pd.read_excel(file_path, engine=engine)
            df.columns = df.columns.str.strip()
            return df
        except ImportError as ie:
            raise ValueError(f"Missing required Excel engine: {str(ie)}")
        except Exception as e:
            raise ValueError(f"Failed to read Excel file: {str(e)}")
    #function to process the daily report file
    def process_daily_report(self, file, filename):
        if not file:
            raise ValueError("No daily report file provided.")

        file_path = os.path.join(self.upload_folder, filename)
        file.save(file_path)
        df = self._read_excel(file_path, filename)
        required_columns = ['Date', 'Candidate Name', 'Role', 'Interview', 'Status', 'Remark']
        missing_columns = [col for col in required_columns if col not in df.columns]
        #if any of the required columns are missing, remove the file and raise an error
        if missing_columns:
            os.remove(file_path) 
            raise ValueError(f"Error: Missing columns {missing_columns} in daily report file {filename}")
        team_member = self._extract_team_member_from_filename(filename)
        df['Team Member'] = team_member
        return df, team_member, file_path
    
    #function to process the new employee file
    def process_new_employee_file(self, file, filename):
        if not file:
            raise ValueError("No new employee file provided.")

        file_path = os.path.join(self.upload_folder, filename)
        file.save(file_path)

        df = self._read_excel(file_path, filename)
        # Check for required columns in the new employee file
        required_columns = ['Employee Name', 'Join Date', 'Role']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            os.remove(file_path)
            raise ValueError(f"Error: Missing columns {missing_columns} in new employee file {filename}")
        return df, file_path
    #function to extract the team member name from the filename
    def _extract_team_member_from_filename(self, filename):
        pattern = r"Daily[ _]?report_(\d+)_([A-Za-z]+)_([A-Za-z]+)\.(?:xls|xlsx)"
        match = re.search(pattern, filename, re.IGNORECASE)
        if match:
            return f"{match.group(2)} {match.group(3)}"
        return "Unknown Team Member"