from flask import Flask, request, send_file, Response
from flask_cors import CORS
import io
import requests
import pandas as pd
import concurrent.futures

app = Flask(__name__)
CORS(app)

@app.route('/generate_report', methods=['POST'])
def generate_report():
    start_date = request.json['start_date']
    end_date = request.json['end_date']
    

# Define the base URL and headers
    base_url = "https://reports.idc.net.pk/metacubes_service/api/BussinessSuite/GetBranchWiseVisitCountAnalytics"
    headers = {"Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJleHAiOjE2NjY1NTQ5NjUsImlzcyI6ImhlbGxvLmJlbGxvLmlzc3VlciIsImF1ZCI6ImhlbGxvLmJlbGxvLmF1ZGllbmNlIn0.hzWYaMV-BeDWBXp1TKn9x25i7uX2-iWPM2QVmnqhyVw", 
            "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Mobile"}

    # Define the payloads for each report
    payloads = [
        {"SubSectionIDs": "1,51,61,55,2,3,4,5,47,6,7,9,10,58,45,39,11,46,37,43,56,12,36,13,14,15,50,57,52,53,42,59,16,17,18,44,34,35,49,30,41,38,20,21,22,60,23,54,48,24,25,26,28,29", "FilterBy": "2", "LabDeptID": -1},
        {"SubSectionIDs": "1,61,55,2,3,4,5,6,9,10,11,56,13,14,15,57,52,53,42,59,16,17,30,41,20,21,22,60,23,54,24,26,28", "FilterBy": "2", "LabDeptID": 1},
        {"SubSectionIDs": "7", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "18", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "47", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "39", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "25", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "29", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "35", "FilterBy": "3", "LabDeptID": 2},
        {"SubSectionIDs": "45", "FilterBy": "3", "LabDeptID": 2}
    ]

    # Create a session for each report
    sessions = [requests.Session() for _ in range(len(payloads))]

    # Define a function to make requests for each report
    def get_report_data(payload, session):
        payload["DateFrom"] = start_date
        payload["DateTo"] = end_date
        payload["GroupBy"] = 1
        payload["LocIDs"] = "-1"
        payload["TPIDs"] = ""
        response = session.post(base_url, headers=headers, json=payload)
        response.raise_for_status()
        data = response.json()["PayLoad"]
        df = pd.DataFrame(data)
        return df

    # Make a request for each report using parallel processing
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        for i, (payload, session) in enumerate(zip(payloads, sessions)):
            payload["DateFrom"] = start_date
            payload["DateTo"] = end_date
            payload["GroupBy"] = 1
            payload["LocIDs"] = "-1"
            payload["TPIDs"] = ""
            future = executor.submit(session.post, base_url, headers=headers, json=payload)
            futures.append((future, i+1))

        dfs = []
        for future, report_num in futures:
            response = future.result()
            response.raise_for_status()
            data = response.json()["PayLoad"]
            df = pd.DataFrame(data)
            df["Report"] = report_num
            dfs.append(df)

    # Concatenate the results and write to Excel file
    results = pd.concat(dfs, ignore_index=True)
    # create a dictionary to replace values in the 'Report' column
    report_dict = {'1': 'Revenue_Ex_PCR', '2': 'Lab_Ex_PCR', '3': 'CT', '4': 'MRI', '5': 'Contrast', '6': 'Doppler', '7': 'USG', '8': 'XRAY', '9': 'OPG', '10': 'DEXA'}
    result = results.copy()
    result['Report'] = result['Report'].astype(str).replace(report_dict)
    result.loc[result['TPCode'].notnull(), 'BranchName'] = result['TPCode']
    result = result.drop(columns=['TPCode']).fillna(0)
    column_list = list(result)
    column_list.remove("BranchName")
    column_list.remove("LocID")
    column_list.remove("Report")

    sheets = {
        'Revenue_Ex_PCR': result[result['Report'] == 'Revenue_Ex_PCR'],
        'Lab_Ex_PCR': result[result['Report'] == 'Lab_Ex_PCR'],
        'CT': result[result['Report'] == 'CT'],
        'MRI': result[result['Report'] == 'MRI'],
        'Contrast': result[result['Report'] == 'Contrast'],
        'Doppler': result[result['Report'] == 'Doppler'],
        'USG': result[result['Report'] == 'USG'],
        'XRAY': result[result['Report'] == 'XRAY'],
        'OPG': result[result['Report'] == 'OPG'],
        'DEXA': result[result['Report'] == 'DEXA']
    }

    # Your existing code to generate the report goes here
    # Make sure to replace the 'input' function with the variables above
    # ...

    # Create an ExcelWriter object and write each DataFrame to a separate sheet
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in sheets.items():
            sheet_data = sheet_data.copy()
            sheet_data[sheet_name] = sheet_data[column_list].sum(axis=1)
            sheet_data = sheet_data.drop(columns=['Report', 'LocID'])
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    response = Response(output.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers.set('Content-Disposition', 'attachment', filename='output.xlsx')
    return response

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)