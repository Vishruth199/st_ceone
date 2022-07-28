# from tempfile import tempdir
import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
import xlwings as xw
import pythoncom
import regex as re
import os
from streamlit_option_menu import option_menu
from PIL import Image
import zipfile
import hydralit_components as hc

import datetime


def upload_single_file(uploaded_file, dir_path):

    with open(os.path.join(dir_path, uploaded_file.name), 'wb') as f:

        f.write(uploaded_file.getbuffer())
        # print()
        # return st.success('saved file : {} in Directory'.format(uploaded_file.name))


def upload_multiple_file(uploaded_file, multiple_dir_path):

    with open(os.path.join('multipleFiles', uploaded_file.name), 'wb') as f:
        f.write(uploaded_file.getbuffer())
        print()
        return st.success('saved file : {} in Directory'.format(uploaded_file.name))


def transform_uploaded_file(dir_path_for_all):
    pythoncom.CoInitialize()

    excel_file_list = []
    for path in os.listdir(dir_path_for_all):
        # check if current path is a file
        if os.path.isfile(os.path.join(dir_path_for_all, path)):
            excel_file_list.append(path)
    print(excel_file_list)

    for eachExcel_file in excel_file_list:
        workbook = load_workbook(dir_path_for_all+"/"+eachExcel_file,
                                 data_only=False, read_only=False)
        sheet = workbook["SA-6239-ENG"]
        cell_value = sheet["g57"].value
    # print(cell_value)

# ---------------write extracted cell value to text file----------------

        with open("sample.txt", "w+") as outfile:
            outfile.write(cell_value)

# ----------------------read text file to dataframe------------------------
        dataframe = pd.read_table(
            r"sample.txt", encoding="latin1", header=None)
        df = dataframe[(dataframe[0].str.contains(':|-'))]
        df[0] = df[0].str[3:]
        df = df[0].str.split(':|-', 1, expand=True)
        print(df)

        print("\n________\n")

# ------------------load workobook with xlwings module --------------------
        app = xw.App(visible=False)
        workbook = xw.Book(dir_path_for_all+"/"+eachExcel_file)
        sheet = workbook.sheets['SA-6239-ENG']

# first column of dataframe
        sheet.range('G15').options(index=False, header=False).value = df[0]
        sheet.range('Z15').options(
            index=False, header=False).value = df[1]  # second column
        workbook.save(dir_path_for_all+"/"+eachExcel_file)
        workbook.close()


def transform_mto_file(uploadedfile):
    df = pd.read_excel("singleFile/"+uploadedfile)
    print(df)
    df = df.drop(['Unnamed: 6', 'Unnamed: 7',
                  'Unnamed: 8', 'Unnamed: 1'], axis=1)
    df["Part No."].fillna(method='ffill', inplace=True)
    df['Quantity'] = df['Quantity'].replace({'m': ''}, regex=True)
    df['Quantity'] = df['Quantity'].astype(float, errors='raise')
# df1 = df.groupby(['GType','Component Description']).apply(lambda a: a[:])
# df1

    df2 = df.groupby(['GType', 'Component Description'])['Quantity'].sum()
    dataframe = df2.to_frame()
    dataframe['Contigency'] = "5%"
# print(df2)
    df = dataframe.loc[['PIP', 'TEE', 'RED', 'FLG', 'VALV', 'BOLT', 'CPL',
                        'GAS', 'UNN', '45L', '90L'], :]  # arranged in following order
# print(df)
# df.insert(0, 'Groups',['A','C','D','F','K'])
    df['Contingency_Qty'] = df.apply(lambda row: row.Quantity * 0.05, axis=1)
    df['Total + Contingency'] = df.apply(
        lambda row: row.Quantity + row.Contingency_Qty, axis=1)

    unique = {'PIP': 'A', 'TEE': 'B', 'RED': 'C', 'FLG': 'D', 'VALV': 'E',
              'BOLT': 'F', 'CPL': 'G', 'GAS': 'H', 'UNN': 'I', '45L': 'J', '90L': 'K'}
    item_no = []
    count = 0
    for i in unique:
        inc = 1
        while i == df.index.values[count][0]:
            item_no.append(unique[i]+'.'+str(inc))
            inc += 1
            count += 1
            if count > 38:
                break

    groups = []
    count = 0
    for i in unique:
        inc = 1
        while i == df.index.values[count][0]:
            groups.append(unique[i])
            inc += 1
            count += 1
            if count > 38:
                break

    # df.reset_index(inplace=True)
    # df.insert(0, "Item Number", item_no)
    # df.insert(0, "Group", groups)

    # df2 = df.groupby(['Group','Item Number']).first()
    df.reset_index(inplace=True)
    df.insert(0, "Groups_num", item_no)
    df.insert(0, "Groups", groups)

    new2 = pd.Series([])
    for i in range(len(df)):
        if df["GType"][i] == "VALV":
            new2[i] = '4311-M2TY-5-14-0002'
        else:
            new2[i] = '4311-M2TY-5-14-0004'

    df.insert(4, "Specification Document Reference", new2)

    df.insert(5, "Datasheet Document/Piping Material Class Specification Reference",
              '4311-M2TY-5-14-0001')

    new = pd.Series([])
    for i in range(len(df)):
        if df["GType"][i] == "PIP":
            new[i] = '6'
        else:
            new[i] = '1'

    df.insert(10, 'Roundup', new)
    df.insert(11, 'Previous Mto Qty', 1)

    new1 = pd.Series([])
    for i in range(len(df)):
        if df["GType"][i] == "PIP":
            new1[i] = 'mtr'
        else:
            new1[i] = 'nos'

    df.insert(12, 'UOM', new1)
    df.insert(13, 'Certification ', '3.1 Certification')

    writer = pd.ExcelWriter(uploadedfile, engine='xlsxwriter')

    df.to_excel(writer, sheet_name='Sheet1',
                startrow=0, header=True, index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    header_format = workbook.add_format({
        'bold': True,
        'fg_color': 'yellow',
        'border': 1})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

        column_len = df[value].astype(str).str.len().max()

        column_len = max(column_len, len(value)) + 3

        worksheet.set_column(col_num, col_num, column_len)

    writer.save()
    with open(uploadedfile, 'rb') as my_file:
        st.download_button(label='Download', data=my_file, file_name=uploadedfile,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# -------------------- DOWNLOADING file------------------------

def download_single_file(directory_path):
    excel_file = os.listdir(directory_path)
    for filename in excel_file:
        with open(directory_path+"/"+filename, 'rb') as my_file:
            st.download_button(label='Download', data=my_file, file_name=filename,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ------------------ ZIPPING multiple files ----------------------

def download_multiple_file(multiple_dir_path, ziph,filename):
    excel_files = os.listdir(multiple_dir_path)

    # basename = "Batch_"
    # suffix = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
    # filename = "_".join([basename, suffix])  # e.g. 'mylogfile_120508_171442'
    # print(filename)
    for root, dirs, files in os.walk(multiple_dir_path):
        for file in excel_files:
            ziph.write(os.path.join(root, file))

    with open(filename+".zip", "rb") as fp:
        st.download_button(label="Download ZIP", data=fp,
                           file_name=filename+".zip", mime="application/zip")


def main():

    with st.sidebar.container():
        image = Image.open("./logo22.png")
        st.image(image, use_column_width=True)

    dir_path = "singleFile"
    multiple_dir_path = "multipleFiles"
    #
    with st.sidebar:

        app_mode = option_menu("NAVIGATION", ["DATA SHEET", "MTO", "TEMPLATE AUTOMATION"],
                               icons=['file-earmark-binary', 'file-earmark-spreadsheet',
                                      'file-spreadsheet-fill'],
                               menu_icon="list", default_index=0,
                               styles={
            "container": {"padding": "!important", "background-color": "#f0ff6"},
            "icon": {"color": "white", "font-size": "28px"},
            "nav-link": {"font-size": "16px", "text-align": "left", "margin": "0px", "--hover-color": "#eeee"},
            "nav-link-selected": {"background-color": "#2C845"},
        }
        )
    if app_mode == "DATA SHEET":
        st.title("Data Sheet")
        menu = ['single excel file trasformation',
                'multiple excel trasformation']
        choice = st.radio('CHOOSE SINGLE OR MULTIPLE FILES', menu)
        st.write(
            '<style>div.row-widget.stRadio > div{flex-direction:row;}</style>', unsafe_allow_html=True)

        if choice == 'single excel file trasformation':
            st.header("Upload single excel files")
            for f in os.listdir(dir_path):
                print("the remaining files are", f)
                os.remove(os.path.join("./singleFile", f))
            datafile = st.file_uploader(
                "upload .xlsx files only", type=['xlsx'])
            print(datafile)
            if datafile is not None:
                st.write("filename:", datafile.name)

            # function to upload and transform
                upload_single_file(datafile, dir_path)

            transform_uploaded_file(dir_path)
            download_single_file(dir_path)

        else:
            if choice == 'multiple excel trasformation':
                st.header("Upload multiple files")
                for f in os.listdir(multiple_dir_path):
                    print("the remaining files are", f)
                    os.remove(os.path.join("./multipleFiles", f))
                datafile = []
                datafile = st.file_uploader(
                    "upload xlsx", type=['xlsx'], accept_multiple_files=True)
                if datafile is not None:
                    for uploaded_file in datafile:
                        upload_multiple_file(uploaded_file, multiple_dir_path)
                    transform_uploaded_file(multiple_dir_path)

                    basename = "Batch_"
                    suffix = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
                    filename = "_".join([basename, suffix])
                    zipf = zipfile.ZipFile(
                        filename + '.zip', 'w', zipfile.ZIP_DEFLATED)
                download_multiple_file(multiple_dir_path, zipf,filename)

                zipf.close()
    if app_mode == "MTO":
        st.title("MTO operation")
        menu = ['single excel file trasformation',
                'multiple excel trasformation']
        choice = st.radio('CHOOSE SINGLE OR MULTIPLE FILES', menu)
        st.write(
            '<style>div.row-widget.stRadio > div{flex-direction:row;}</style>', unsafe_allow_html=True)
        if choice == 'single excel file trasformation':
            st.subheader("Upload files")
            # st.text("text")
            for f in os.listdir(dir_path):
                print("the remaining files are", f)
                os.remove(os.path.join("./singleFile", f))
            datafile = st.file_uploader(
                "upload xlsx", type=['xlsx'], accept_multiple_files=False)
            if datafile is not None:
                st.write("filename:", datafile.name)
                print(datafile.name)
                upload_single_file(datafile, dir_path)

            # upload_one_file(datafile)
        # input_file="single_folder"
                transform_mto_file(datafile.name)
            # download_one_file(input_file)

        else:
            st.subheader('About')


if __name__ == '__main__':
    main()
