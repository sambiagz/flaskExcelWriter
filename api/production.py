import os
import pandas as pd
from xlsxwriter import Workbook
from werkzeug.utils import secure_filename

file_name_2 = "results.xlsx"
workbook = Workbook(file_name_2)
bold = workbook.add_format({'bold': True})
worksheet = workbook.add_worksheet()


def create_row(file_name):
    df = pd.read_excel(file_name, sheet_name="DataBase")
    if(df["PipeNumber"][0]):
        print("Generating Row")
        columns = df.columns.append(df.columns)
        data = df.loc[0].append(df.loc[len(df)-1])
        return data


# Allowed extensions
ALLOWED_EXTENSIONS = set(['csv', 'xlsx', 'xlsm'])
# Get current path
path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')


WELDING_ROWS = []


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def write_file(files):
    for file in files:
        if file and allowed_file(file.filename):
            print("File Found In List")
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                WELDING_ROWS.append(create_row(file_path))
            except:
                print("Error Occured: File is Production")

            df = pd.read_excel(file_path, sheet_name="DataBase")
            columns = df.columns.append(df.columns)

            # WRITE HEADERS
            for i in range(0, len(columns)):
                try:
                    if("Unnamed" in columns[i]):
                        pass
                    else:
                        worksheet.write(0, i, columns[i], bold)
                        worksheet.set_column(0, i, 10)
                except:
                    pass

            # WRITE CONTENT
            for i in range(0, len(WELDING_ROWS)+1):
                for j in range(0, len(columns)):
                    try:
                        worksheet.write((i+1), j, WELDING_ROWS[i][j])
                    except:
                        pass
        else:
            print("FILE NOT IN LIST")
    workbook.close()
