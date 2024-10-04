from flask import Flask, request, send_file, render_template
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Get the uploaded files from the form
        file1 = request.files['file1']
        file2 = request.files['file2']

        # Load the Excel files into pandas DataFrames
        df_file1 = pd.read_excel(file1, dtype=str, keep_default_na=False)
        df_file2 = pd.read_excel(file2, dtype=str, keep_default_na=False)
        print("Sheets loaded successfully.")

        # Filter and find the matching rows
        df_file1_unique = df_file1[(df_file1['Empty Perfects'] != '1') & (df_file1['Perfect with External Returns'] != '1')]
        similar_rows = []

        for i, row_df1 in df_file1_unique.iterrows():
            for j, row_df2 in df_file2.iterrows():
                # Compare specific columns and calculate similarity
                if ((row_df1['Site'] == row_df2[' Site ']) and
                    (row_df1['Perfects Project'] == row_df2[' Project ']) and
                    (abs(float(row_df2[' Total Net Time TM with Empty clips (HRS) ']) - float(row_df1['Net Time AIP [h]'])) < 1)):
                    similar_rows.append(row_df1)

        # Convert the similar rows to a DataFrame
        similar_df = pd.DataFrame(similar_rows)

        # Save the results to an Excel file
        output_file = 'similar_rows_output.xlsx'
        with pd.ExcelWriter(output_file) as writer:
            similar_df.to_excel(writer, sheet_name="SimilarRows", index=False)

        # Send the Excel file to be downloaded
        return send_file(output_file, as_attachment=True)

    except Exception as e:
        return {'error': str(e)}, 500

if __name__ == '__main__':
    app.run(debug=True)
