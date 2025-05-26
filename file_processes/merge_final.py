import pandas as pd
from datetime import datetime
import os
class FinalReportGenerator:
    def __init__(self, output_folder):
        self.output_folder = output_folder

    def generate_dashboard(self, all_interviews_df, new_employees_df):
        #Generate dashboard with merged data for passing candidates
        pass_interviews = all_interviews_df[all_interviews_df['Status'] == 'Pass'].copy()

        # Perform the merge operation
        dashboard = pd.merge(
            new_employees_df[['Employee Name', 'Join Date', 'Role']],
            pass_interviews[['Candidate Name', 'Team Member']],
            how='left',
            left_on='Employee Name',
            right_on='Candidate Name'
        )

        # Select and rename columns for the final result
        result = dashboard[['Employee Name', 'Join Date', 'Role', 'Team Member']].copy()

        # Convert Join Date to datetime and format
        try:
            result['Join Date'] = pd.to_datetime(result['Join Date']).dt.strftime('%d-%b-%Y')
        except Exception as e:
            raise ValueError(f"Error formatting Join Date: {str(e)}")

        return result

    def assign_team_member_for_unmatched(self, result, available_team_members):
        number_of_team_members = len(available_team_members)
        if number_of_team_members == 0:
            return # No team members to assign
        # Identify rows where 'Team Member' is None
        unmatched_data = result[result['Team Member'].isna()].index
        for i, index in enumerate(unmatched_data):
            team_member = available_team_members[i % number_of_team_members]
            result.at[index, 'Team Member'] = team_member


    def save_report(self, df):
        # Save the DataFrame to an Excel file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"dashboard_{timestamp}.xlsx"
        output_filepath = os.path.join(self.output_folder, output_filename)

        try:
            df.to_excel(output_filepath, index=False, engine='openpyxl')
            return output_filepath
        except Exception as e:
            raise ValueError(f"Error saving report to {output_filepath}: {str(e)}")