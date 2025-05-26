from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
from file_processes.file_processing import FileProcessor
from file_processes.merge_final import FinalReportGenerator
import pandas as pd
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'Uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx'}

#Create directories for uploads and output folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

#Function to check the extension of the uploded file
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

#Main route for the Flask app
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        errors = []

        # Initialize file processor and report generator
        file_processor = FileProcessor(app.config['UPLOAD_FOLDER'])
        report_generator = FinalReportGenerator(app.config['OUTPUT_FOLDER'])

        # Get files from input via GET method
        daily_report1 = request.files.get('daily_report1')
        daily_report2 = request.files.get('daily_report2')
        new_employee = request.files.get('new_employee_file')

        # Validate inputs for empty files
        if not daily_report1 or daily_report1.filename == '':
            errors.append("Please upload the first daily report file.")
        if not daily_report2 or daily_report2.filename == '':
            errors.append("Please upload the second daily report file.")
        if not new_employee or new_employee.filename == '':
            errors.append("Please upload the new employee file.")
        if errors:
            return render_template('error.html', errors=errors)

        uploaded_file_paths = []
        interviews = []
        extracted_team_members = []

        try:
            # Process daily report 1
            if allowed_file(daily_report1.filename):
                filename1 = secure_filename(daily_report1.filename)
                interview1, team_member1, path1 = file_processor.process_daily_report(daily_report1, filename1)
                interviews.append(interview1)
                extracted_team_members.append(team_member1)
                uploaded_file_paths.append(path1)
            else:
                errors.append(f"Invalid file type for {daily_report1.filename}")

            # Process daily report 2
            if allowed_file(daily_report2.filename):
                filename2 = secure_filename(daily_report2.filename)
                interview2, team_member2, path2 = file_processor.process_daily_report(daily_report2, filename2)
                interviews.append(interview2)
                extracted_team_members.append(team_member2)
                uploaded_file_paths.append(path2)
            else:
                errors.append(f"Invalid file type for {daily_report2.filename}")

            if not interviews:
                errors.append("No valid daily report files processed.")
                return render_template('error.html', errors=errors)

            all_interviews= pd.concat(interviews, ignore_index=True)

            # Process new employee file
            if allowed_file(new_employee.filename):
                new_employee_filename = secure_filename(new_employee.filename)
                new_employees_df, new_employee_path = file_processor.process_new_employee_file(new_employee, new_employee_filename)
                uploaded_file_paths.append(new_employee_path)
            else:
                errors.append(f"Invalid file type for {new_employee.filename}")

            # Generate dashboard with merged data
            dashboard= report_generator.generate_dashboard(all_interviews, new_employees_df)

            # Assign team members for unmatched cases
            report_generator.assign_team_member_for_unmatched(dashboard, extracted_team_members)

            # Save report
            output_filepath = report_generator.save_report(dashboard)

            # Clean up uploaded files
            for path in uploaded_file_paths:
                if os.path.exists(path):
                    os.remove(path)

            if os.path.exists(output_filepath):
                return send_file(
                    output_filepath,
                    as_attachment=True,
                    download_name=os.path.basename(output_filepath)
                )
            else:
                errors.append("Error: Failed to create output file.")

        except ValueError as e:
            errors.append(str(e))
        except Exception as e:
            errors.append(f"An unexpected error occurred")
        finally:
            for path in uploaded_file_paths:
                if os.path.exists(path):
                    os.remove(path)
        return render_template('errors.html', errors=errors)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)