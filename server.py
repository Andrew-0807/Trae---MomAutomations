from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import traceback
import zipfile  # Add this import

try:
    # Import the processor modules
    from classes.valoare_sgr import SGRValueProcessor
    from classes.valoare_minus import ValoareMinus
    from classes.format_add_column import FormatAddColumn
    from classes.excel_data_extractor import ExcelDataExtractor
except Exception as e:
    print(f"Error importing modules: {str(e)}")
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    files = request.files.getlist('file')  # Get all uploaded files
    process_type = request.form['process_type']
    
    try:
        outputs = []
        filenames = []
        for file in files:
            # Check if the file has a valid Excel extension
            if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                print(f"Skipping non-Excel file: {file.filename}")
                continue
            # Check if the file is not empty
            file.seek(0, io.SEEK_END)
            file_length = file.tell()
            file.seek(0)
            if file_length == 0:
                print(f"Skipping empty file: {file.filename}")
                continue

            try:
                df = pd.read_excel(file, engine='openpyxl')
                df.name = file.filename
            
                # Process the data based on the process_type
                if process_type == 'adaos':
                    processor = FormatAddColumn()
                elif process_type == 'sgr':
                    processor = SGRValueProcessor()
                elif process_type == 'minus':
                    processor = ValoareMinus()
                elif process_type == 'extract':
                    processor = ExcelDataExtractor()
                else:
                    return "Invalid process type", 400
                
                # Process the data
                result_df = processor.process_dataframe(df)
                
                # Save the processed DataFrame to a BytesIO object
                output = io.BytesIO()
                result_df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                outputs.append(output)
                # Save the filename for the zip
                original_filename = file.filename
                processed_filename = f"{process_type} - {original_filename}"
                filenames.append(processed_filename)
            except Exception as e:
                print(f"Error reading {file.filename}: {e}")
                continue

        # These lines should be OUTSIDE the for loop!
        if len(outputs) == 1:
            return send_file(outputs[0], download_name=filenames[0], as_attachment=True)
        
        # If multiple files, zip them
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            for output, fname in zip(outputs, filenames):
                output.seek(0)
                zipf.writestr(fname, output.read())
        zip_buffer.seek(0)
        return send_file(zip_buffer, download_name="processed_files.zip", as_attachment=True, mimetype='application/zip')
        
    except Exception as e:
        traceback.print_exc()
        return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')