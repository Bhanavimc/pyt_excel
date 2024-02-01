from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import subprocess

app = Flask(__name__)

def open_excel_file():
    try:
        subprocess.Popen(['start', 'datas.xlsx'], shell=True)
    except Exception as e:
        print(f"Error opening Excel file: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add', methods=['POST'])
def add_data():
    name = request.form.get('name')
    usn = request.form.get('usn')
    branch = request.form.get('branch')
    sec = request.form.get('sec')
    hobbies = request.form.get('hobbies')

    new_data = pd.DataFrame({'Name': [name], 'USN': [usn], 'Branch': [branch], 'Sec': [sec], 'Hobbies': [hobbies]})

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    except FileNotFoundError:
        updated_data = new_data

    updated_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
    open_excel_file()
    return render_template('index.html', table=updated_data.to_html(classes='table table-bordered table-striped', index=False))

@app.route('/delete', methods=['POST'])
def delete_data():
    name = request.form.get('delete_name')

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        updated_data = existing_data[existing_data['Name'] != name]
    except FileNotFoundError:
        updated_data = pd.DataFrame()

    updated_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
    open_excel_file()
    return render_template('index.html', table=updated_data.to_html(classes='table table-bordered table-striped', index=False))

if __name__ == '__main__':
    app.run(debug=True)

