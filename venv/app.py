from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os

app = Flask(__name__)

# Use absolute path for the Excel file
DATA_FILE = os.path.join(os.getcwd(), 'examiner_data.xlsx')

# Initialize Excel file if it doesn't exist
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=[
        "Examiner Name", "Mobile No", "Subject", "Date of Passing",
        "Amount of Remuneration", "Travelling Allowance", "Remarks"
    ])
    df.to_excel(DATA_FILE, index=False)

@app.route('/')
def index():
    df = pd.read_excel(DATA_FILE)
    all_entries = df.to_dict(orient='records')
    return render_template("form.html", all_entries=all_entries)

@app.route('/add', methods=['POST'])
def add_entry():
    new_entry = {
        "Examiner Name": request.form['examiner_name'],
        "Mobile No": request.form['mobile_no'],
        "Subject": request.form['subject'],
        "Date of Passing": request.form['date_of_passing'],
        "Amount of Remuneration": request.form['amount_of_remuneration'],
        "Travelling Allowance": ', '.join(request.form.getlist('travelling_allowance')),
        "Remarks": request.form['remarks']
    }

    df = pd.read_excel(DATA_FILE)
    new_df = pd.DataFrame([new_entry])
    df = pd.concat([df, new_df], ignore_index=True)
    df.to_excel(DATA_FILE, index=False)

    return redirect('/')

@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit_entry(id):
    df = pd.read_excel(DATA_FILE)

    if id >= len(df):
        return "Entry does not exist.", 404

    if request.method == 'GET':
        entry = df.iloc[id]
        return render_template('edit.html', entry=entry, id=id)
    else:
        df.at[id, 'Examiner Name'] = request.form['examiner_name']
        df.at[id, 'Mobile No'] = request.form['mobile_no']
        df.at[id, 'Subject'] = request.form['subject']
        df.at[id, 'Date of Passing'] = request.form['date_of_passing']
        df.at[id, 'Amount of Remuneration'] = request.form['amount_of_remuneration']
        df.at[id, 'Travelling Allowance'] = ', '.join(request.form.getlist('travelling_allowance'))
        df.at[id, 'Remarks'] = request.form['remarks']

        df.to_excel(DATA_FILE, index=False)
        return redirect('/')

@app.route('/delete/<int:id>')
def delete_entry(id):
    df = pd.read_excel(DATA_FILE)

    if id >= len(df):
        return "Entry does not exist.", 404

    df.drop(index=id, inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.to_excel(DATA_FILE, index=False)
    return redirect('/')

@app.route('/download')
def download():
    return send_file(DATA_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
