from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/compare', methods=['POST'])
def compare_files():
    # Check if both files are uploaded
    if 'file1' not in request.files or 'file2' not in request.files:
        return "Please upload both files", 400
    
    # Read the uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    try:
        # Load the Excel sheets
        sheet1 = pd.read_excel(file1)
        sheet2 = pd.read_excel(file2)
    except Exception as e:
        return f"Error reading files: {e}", 400
    
    # Print the columns for debugging
    print("Columns in file1:", sheet1.columns.tolist())
    print("Columns in file2:", sheet2.columns.tolist())

    # Ensure the relevant columns exist in each sheet
    if 'Project' in sheet1.columns and 'Total Net Time TM with Empty clips (HRS)' in sheet1.columns and \
       'Perfects Project' in sheet2.columns and 'Net Time AIP [h]' in sheet2.columns:
        
        # Compare the 'Project' from sheet1 with 'Perfects Project' from sheet2
        comparison = pd.merge(sheet1[['Project', 'Total Net Time TM with Empty clips (HRS)']], 
                              sheet2[['Perfects Project', 'Net Time AIP [h]']], 
                              left_on='Project', right_on='Perfects Project', how='inner')
        
        # Prepare the result as an Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            comparison.to_excel(writer, index=False)
        
        output.seek(0)
        return send_file(output, attachment_filename='matching_rows.xlsx', as_attachment=True)
    
    return "One or both of the required columns are missing from the Excel sheets", 400

if __name__ == "__main__":
    app.run(debug=True)
