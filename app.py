# Imports
import os  # Website
import json  # File for Martian Equation
import pandas as pd  # Reading through the excel file
from flask import Flask, request, render_template, redirect  # Needed for website
from werkzeug.utils import secure_filename  # Identifying the factor from the chart

# Declaring the flask, configurations for the website
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'  # Saves for uploads to be used
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}  # Allows for excel files

# Function for obtaining the file, finds out if it is an excel file
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Function for calculating the m3LDLC function
def m3SLDLC(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing value: {', '.join(missing_values)}"
    if TC > 1000 or TG > 3000:
        raise ValueError("TC must be <= 1000 mg/dL and TG must be <= 3000 mg/dL for valid m3SLDL calculation.")
    non_HDL = TC - HDL
    equation = (0.9028 * TC) + (-0.8573 * HDL) + (-0.1042 * TG) + (-0.000472 * (TG * non_HDL)) + (0.0000623 * (TG * TG)) + (pow(non_HDL, 2) * 0.0002866) + 3.0377

    if TG > 800:
        warning = "Warning: LDLC results with TG > 800 mg/dL may have > 30 mg/dL error."
        result = round(equation, 1)
        return f"{result} {warning}"
    if equation < 0:
        return 0.0
    if equation > non_HDL:
        return round(non_HDL, 1)
    return round(equation, 1)

# Function for calculating the SLDLC
def SLDLC(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing value: {', '.join(missing_values)}"
    if TC > 1000 or TG > 1500:
        raise ValueError("TC must be <= 1000 mg/dL and TG must be <= 1500 mg/dL for valid SLDLC calculation.")
    non_HDL = TC - HDL
    equation = (TC * 1.055) - (HDL * 1.029) - ((TG * 0.1168) + ((TG * non_HDL) * 0.000467) - (pow(TG, 2) * 0.00006199)) - 9.4386
    if equation < 0:
        return 0.0
    if equation > non_HDL:
        return round(non_HDL, 1)
    return round(equation, 1)

# Function for calculating eS_LDL
def eS_LDL(TC, HDL, TG, ApoB):
    if any(pd.isna(var) for var in [TC, HDL, TG, ApoB]):
        return ''
    if TC > 1000 or TG > 1500:
        raise ValueError("TC must be <= 1000 mg/dL and TG must be <= 1500 mg/dL for valid enhanced eS_LDL calculation.")
    if HDL > TC:  # Adding a check to ensure non-HDL is not negative
        raise ValueError("HDL must be <= TC for a valid calculation.")
    non_HDL = TC - HDL
    equation = (0.8708 * TC) + (-0.8022 * HDL) + (-0.1432 * TG) + (0.2202 * ApoB) + (0.000808 * (TG * ApoB)) + (-0.000896 * (TG * non_HDL)) + ((pow(TG, 2) * 0.000112)) - 4.726
    if equation < 0:
        return 0.0
    if equation > non_HDL:
        return round(non_HDL, 1)
    return round(equation, 1)

# Function for calculating SLDLC in mmol/L
def SLDLC_mmol(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing values: {', '.join(missing_values)}"
    if TC > 25.9 or TG > 16.9:
        raise ValueError("TC must be <= 26 mmol/L and TG must be <= 17 mmol/L for valid SLDLC calculation in mmol/L.")
    non_HDL = TC - HDL
    equation = (TC / 0.948) - (HDL / 0.971) - ((TG / 3.47) + ((TG * non_HDL) / 24.16) - (pow(TG, 2) / 79.36)) - 0.224
    if equation < 0:
        return 0.0
    if equation > non_HDL:
        return round(non_HDL, 1)
    return round(equation, 1)

# Function for calculating the Frienwald Equation
def FLDLC(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing values: {', '.join(missing_values)}"
    if TC > 400:
        raise ValueError("TC must be < 400 mg/dL for valid FLDLC calculation in mg/dL.")
    equation = TC - HDL - (TG / 5)
    return round(equation, 1)

# Function for the Martian Hopkins Equation
def MLDLC(TC, HDL, TG):
    missing_fields = [field for field, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_fields:
        return f"Missing values: {', '.join(missing_fields)}"
    with open('data/data.json', 'r') as json_file:
        data = json.load(json_file)

    factor = None
    non_HDL = TC - HDL
    if TG > 799:
        TG = 799
        rounded_TG = TG
    else:
        rounded_TG = round(TG)
    non_HDL = round(non_HDL)
    for entry in data:
        if entry['TG_low'] <= rounded_TG <= entry['TG_high']:
            for non_hdl_range in entry['nonHDL-C']:
                if non_hdl_range['low'] <= non_HDL <= non_hdl_range['high']:
                    factor = non_hdl_range['factor']
                    break
            break

    if factor is not None:
        equation = (TC - HDL - (TG / factor))
        return round(equation, 1)
    else:
        raise ValueError("Factor not found for the given inputs.")

# Conversion functions for mmol/L
def convert_mg(value, factor):
    if isinstance(value, (int, float)):
        return value * factor
    return None

def m3SLDLC_mmol(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing values: {', '.join(missing_values)}"
    else:
        if TC > 25.9 or TG > 33.9:
            raise ValueError("TC must be <= 25.9 mmol/L and TG must be <= 33.9 mmol/L for valid m3LDLC calculation in mmol/L.")
        TC_mmol = convert_mg(TC, 38.67)
        HDL_mmol = convert_mg(HDL, 38.67)
        TG_mmol = convert_mg(TG, 88.57)
        equation = m3SLDLC(TC_mmol, HDL_mmol, TG_mmol)
        equation = convert_mg(equation, 0.02585983966)
        if TG > 9.0:
            warning = "Warning: LDLC results with TG > 9.0 mg/dL may have > 0.3 mg/dL error."
            result = round(equation, 1)
            return f"{result} {warning}"
    return round(equation, 1)

def eS_LDL_mmol(TC, HDL, TG, ApoB):
    if any(pd.isna(var) for var in [TC, HDL, TG, ApoB]):
        return ''
    else:
        if TC > 25.9 or TG > 16.9:
            raise ValueError("TC must be <= 25.9 mmol/L and TG must be <= 16.9 mmol/L for valid enhanced eS_LDL calculation.")
        TC_mmol = convert_mg(TC, 38.67)
        HDL_mmol = convert_mg(HDL, 38.67)
        TG_mmol = convert_mg(TG, 88.57)
        ApoB_mmol = convert_mg(ApoB, 100)
        equation = eS_LDL(TC_mmol, HDL_mmol, TG_mmol, ApoB_mmol)
        equation = convert_mg(equation, 0.02585983966)
    return round(equation, 1)

def FLDLC_mmol(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing values: {', '.join(missing_values)}"
    else:
        if TC > 10.3:
            raise ValueError("TC must be < 10.3 mmol/L for valid FLDLC calculation in mmol/L.")
        TC_mmol = convert_mg(TC, 38.67)
        HDL_mmol = convert_mg(HDL, 38.67)
        TG_mmol = convert_mg(TG, 88.57)
        equation = FLDLC(TC_mmol, HDL_mmol, TG_mmol)
        equation = convert_mg(equation, 0.02585983966)
    return round(equation, 1)

def MLDLC_mmol(TC, HDL, TG):
    missing_values = [factor for factor, value in zip(["TC", "HDL", "TG"], [TC, HDL, TG]) if pd.isna(value)]
    if missing_values:
        return f"Missing values: {', '.join(missing_values)}"
    else:
        TC_mmol = convert_mg(TC, 38.67)
        HDL_mmol = convert_mg(HDL, 38.67)
        TG_mmol = convert_mg(TG, 88.57)
        equation = MLDLC(TC_mmol, HDL_mmol, TG_mmol)
        equation = convert_mg(equation, 0.02585983966)
    return round(equation, 1)

@app.route('/')  # Used to create a website link
def upload_form():
    return render_template('upload.html')  # The file is uploaded to the upload.html file

@app.route('/', methods=['POST'])  # Post the website
def upload_file():
    if 'file' not in request.files:  # Checks if the file is there
        return redirect(request.url)  # website url
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    unit = request.form.get('unit')  # Get the selected unit from the form
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)  # Website - upload file button
        file.save(filepath)
        results = process_file(filepath, unit)  # Pass the selected unit to the process_file function
        return render_template('results.html', results=results, unit=unit)  # Page with results and unit
    return redirect(request.url)

def process_file(filepath, unit):
    df = pd.read_excel(filepath)  # Reading the excel file
    df.columns = df.columns.str.strip()  # Stripping whitespace from column names
    print(df.columns)
    results = []  # Empty Results List

    for index, row in df.iterrows():  # Loop through the rows of the dataframe
        patient_id = str(int(row.get('PatientID')))
        HDL = row.get('HDLC')
        TC = row.get('TC')
        TG = row.get('TG')
        ApoB = row.get('ApoB')

        result = {'PatientID': patient_id}  # Store patient ID

        if unit == 'mg/dl':
            try:
                result['SLDLC'] = SLDLC(TC, HDL, TG)
            except ValueError as error:
                result['SLDLC'] = str(error)
            try:
                result['eS_LDL'] = eS_LDL(TC, HDL, TG, ApoB)
            except ValueError as error:
                result['eS_LDL'] = str(error)
            try:
                result['m3SLDLC'] = m3SLDLC(TC, HDL, TG)
            except ValueError as error:
                result['m3SLDLC'] = str(error)
            try:
                result['FLDLC'] = FLDLC(TC, HDL, TG)
            except ValueError as error:
                result['FLDLC'] = str(error)
            try:
                result['MLDLC'] = MLDLC(TC, HDL, TG)
            except ValueError as error:
                result['MLDLC'] = str(error)

        elif unit == 'mmol/l':
            try:
                result['SLDLC_mmol'] = SLDLC_mmol(TC, HDL, TG)
            except ValueError as error:
                result['SLDLC_mmol'] = str(error)
            try:
                result['eS_LDL_mmol'] = eS_LDL_mmol(TC, HDL, TG, ApoB)
            except ValueError as error:
                result['eS_LDL_mmol'] = str(error)
            try:
                result['m3SLDLC_mmol'] = m3SLDLC_mmol(TC, HDL, TG)
            except ValueError as error:
                result['m3SLDLC_mmol'] = str(error)
            try:
                result['FLDLC_mmol'] = FLDLC_mmol(TC, HDL, TG)
            except ValueError as error:
                result['FLDLC_mmol'] = str(error)
            try:
                result['MLDLC_mmol'] = MLDLC_mmol(TC, HDL, TG)
            except ValueError as error:
                result['MLDLC_mmol'] = str(error)

        results.append(result)  # Add the result to the list

    return results

if __name__ == "__main__":
    app.run(debug=True)
