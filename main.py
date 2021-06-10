import streamlit as st
import os,zipfile,shutil
from io import StringIO
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import requests,time,pythoncom
import base64
from mailmerge import MailMerge

invoice = os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoices/")
proforma_temp = os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma Template/")
other_doc = os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Other Documents/")
#Title
favicon = "logo.png"
st.set_page_config(page_title='Garibsons Pvt. Ltd.', layout = 'centered', page_icon = favicon, initial_sidebar_state = 'expanded') #page_icon = favicon,
# st.sidebar.title("Garibsons Pvt. Ltd.")
hide_menu_style = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        </style>
        """
st.markdown(hide_menu_style, unsafe_allow_html=True)
st.sidebar.markdown(
    """
<style>
.sidebar .sidebar-content {
    background-image: linear-gradient(#E7E6DD,#808080);
    color: Black;
}
</style>
""",
    unsafe_allow_html=True,
)
st.markdown('<style>body{background-image: logo.jpg ;}</style>',unsafe_allow_html=True)
st.title("Garibsons Pvt. Ltd.")
st.text("Version 2.0")
# Defining Directories
#CLIENT_FOLDER = './Garibson Web App/'
#file_list = os.listdir(CLIENT_FOLDER)
#dic = {key: time.ctime(os.path.getmtime(
    #os.path.join(CLIENT_FOLDER, key))) for key in file_list}
page = st.selectbox("Choose the Form Type", ["Proforma Invoice","Invoice- Other Documents","Custom Invoice | Cutomer Invoice | Bill of Lading"])

#Selection of Page
if page == 'Proforma Invoice':
    st.header("Proforma Invoice")
    st.sidebar.title("Templates")
    st.sidebar.write("Please select Templates")
    side_file = [st.sidebar.checkbox(f, key=f) for f in proforma_temp]
    file_n = [file for file, checked in zip(proforma_temp, side_file) if checked]
    usr_input = st.text_input(label="Proforma Invoice No.", max_chars=20)
    file_name = st.file_uploader(label="Please upload Docx Template",accept_multiple_files=True)
    submit = st.button(label="Submit")

    if submit:
        for file in file_n:
            x = os.path.join("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma Template/",file)
            try:
                doc = DocxTemplate(os.path.join("/Proforma Template/",x))
                x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(usr_input))
                data = x.json()
                for x in data['items']:
                    doc.render(x)
                    doc.save(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file}.docx")
                    pythoncom.CoInitialize()
                    convert(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file}.docx",f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file}.pdf")
            except Exception as e:
                st.warning(e)
        if file_name is not None:
            for file in file_name:
                doc = DocxTemplate(file)
                doc.save(f"{file.name}")
                for x in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/"):
                    # st.write(x)
                    if x.endswith(".docx"):
                        doc = DocxTemplate(x)
                        x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(usr_input))
                        data = x.json()
                        for x in data['items']:
                            doc.render(x)
                            doc.save(
                                f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file.name}.docx")
                            pythoncom.CoInitialize()
                            convert(
                                f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file.name}.docx",
                                f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file.name}.pdf")
        st.success("Process completed.")
        zipf = zipfile.ZipFile(f'{usr_input}_doc.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/'):
            for file in files:
                if file[-4:] == 'docx':
                    zipf.write('./Proforma/' + file)
            zipf.close()
        zipf = zipfile.ZipFile(f'{usr_input}_pdf.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/'):
            for file in files:
                if file[-3:] == 'pdf':
                    zipf.write('./Proforma/' + file)
            zipf.close()
            def get_binary_file_downloader_html(bin_file, file_label='File'):
                with open(bin_file, 'rb') as f:
                    data = f.read()
                bin_str = base64.b64encode(data).decode()
                href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}"><button class="streamlit-button small-button primary-button ">{bin_file}</button></a>'
                return href
            pdf = st.markdown(get_binary_file_downloader_html(f'{usr_input}_pdf.zip', 'download'), unsafe_allow_html=True)
            doc = st.markdown(get_binary_file_downloader_html(f'{usr_input}_doc.zip', 'download'), unsafe_allow_html=True)
            if pdf and doc:
                for file in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/"):
                    os.remove(os.path.join(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{file}"))
            for file in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/"):
                if file.endswith(".docx"):
                    os.remove(os.path.join(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/{file}"))
            btn = st.button("Clear")
            if btn:
                st.cache(func="clear")
# Selection of other page
if page == 'Invoice- Other Documents':
    st.header("Invoice - Other Documents")
    st.sidebar.title("Templates")
    st.sidebar.write("Please select Templates")
    side_file = [st.sidebar.checkbox(f, key=f) for f in other_doc]
    file_n = [file for file, checked in zip(other_doc, side_file) if checked]
    usr_input = st.text_input(label="Invoice No.", max_chars=20)
    file_name = st.file_uploader(label="Please upload Docx Template")
    submit = st.button(label="Submit")
    if submit:
        for file in file_n:
            x = os.path.join(
                "C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Other Documents/",
                file)
            try:
                doc = DocxTemplate(os.path.join(
                    "C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Other Documents/",
                    x))
                x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?invno={}'.format(usr_input))
                data = x.json()
                for x in data['items']:
                    doc.render(x)
                    doc.save(
                        f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/oth_docs/{file}.docx")
                    pythoncom.CoInitialize()
                    convert(
                        f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/oth_docs/{file}.docx",
                        f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/oth_docs/{file}.pdf")
                st.success("Your file is ready to download")
            except Exception as e:
                st.warning(e)
        if file_name is not None:
            for file in file_name:
                doc = DocxTemplate(file)
                doc.save(f"{file.name}")
                for x in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/"):
                    # st.write(x)
                    if x == file.name:
                        if x.endswith(".docx"):
                            doc = DocxTemplate(x)
                            x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(usr_input))
                            data = x.json()
                            for x in data['items']:
                                doc.render(x)
                                doc.save(
                                    f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file.name}.docx")
                                pythoncom.CoInitialize()
                                convert(
                                    f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file.name}.docx",
                                    f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Proforma/{usr_input}_{file.name}.pdf")
        st.success("Process completed.")


        zipf = zipfile.ZipFile('download_doc.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/oth_docs/'):
            for file in files:
                if file[-4:] == 'docx':
                    zipf.write('./oth_docs/' + file)
            zipf.close()
        zipf = zipfile.ZipFile(f'download_pdf.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/oth_docs/'):
            for file in files:
                if file[-3:] == 'pdf':
                    zipf.write('./oth_docs/' + file)
            zipf.close()
            def get_binary_file_downloader_html(bin_file, file_label='File'):
                with open(bin_file, 'rb') as f:
                    data = f.read()
                bin_str = base64.b64encode(data).decode()
                href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}"><button class="streamlit-button small-button primary-button ">{bin_file}</button></a>'
                return href


            pdf = st.markdown(get_binary_file_downloader_html('download_pdf.zip', 'download'), unsafe_allow_html=True)
            doc = st.markdown(get_binary_file_downloader_html('download_doc.zip', 'download'), unsafe_allow_html=True)
            if pdf and doc:
                for file in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/oth_docs/"):
                    os.remove(os.path.join(f"./oth_docs/{file}"))
            for file in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/"):
                if file.endswith(".docx"):
                    os.remove(os.path.join(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/{file}"))
            btn = st.button("Clear")
            if btn:
                st.cache(func="clear")

#----------------------------------------------------------------------------------------------------------
#Invoices
if page == 'Custom Invoice | Cutomer Invoice | Bill of Lading':
    st.header("Custom Invoice | Cutomer Invoice | Bill of Lading")
    st.sidebar.title("Templates")
    st.sidebar.write("Please select Templates")
    side_file = [st.sidebar.checkbox(f, key=f) for f in invoice]
    file_n = [file for file, checked in zip(invoice, side_file) if checked]
    usr_input = st.text_input(label="Invoice No.", max_chars=20)
    file_name = st.file_uploader(label="Please upload Docx Template",accept_multiple_files=True)
    submit = st.button(label="Submit")

    if submit:
        for file in file_n:
            x = os.path.join("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoices/",file)
            try:
                doc = MailMerge(os.path.join("./Invoices/",x))
                data = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/pck_lst?INVNO={}'.format(usr_input))
                data = data.json()
                res = [dict([key, str(value)] for key, value in dicts.items()) for dicts in data['items']]
                inv = [data['invno'] for data in data['items']]
                efrm = [data['e_frmno'] for data in data['items']]
                shipper = [data['text_1'] for data in data['items']]
                shipper_addr = [data['long_text_1'] for data in data['items']]
                notify_name = [data['text_2'] for data in data['items']]
                notify_addr = [data['long_text_2'] for data in data['items']]
                port_load = [data['text_3'] for data in data['items']]
                port_dis = [data['text_4'] for data in data['items']]
                inv_dsc = [data['inv_desc'] for data in data['items']]
                text_1 = [data['text_1'] for data in data['items']]
                text_2 = [data['text_2'] for data in data['items']]
                long_text_1 = [data['long_text_1'] for data in data['items']]
                long_text_2 = [data['long_text_2'] for data in data['items']]
                blno = [data['blno'] for data in data['items']]
                price = [data['unit_price'] for data in data['items']]
                df = pd.DataFrame(data['items'])
                sum_bagqty = df['no_of_bg'].sum()
                sum_bagqty = str(sum_bagqty)
                sum_net_wt_m_tons = df['net_wt_m_tons'].sum()
                sum_net_wt_m_tons = str(sum_net_wt_m_tons)
                sum_gr_dt = df['gr_dt'].sum()
                sum_gr_dt = "{:.3f}".format(sum_gr_dt)
                sum_gr_dt = str(sum_gr_dt)

                cust = {
                            "invno": inv[0],
                            "sum_bagqty":sum_bagqty,
                            "sum_net_wt_m_tons":sum_net_wt_m_tons,
                            "sum_gr_dt":sum_gr_dt,
                            "sum":sum_gr_dt,
                            "eform": efrm[0],
                            "shipper_name":shipper[0],
                            "shipper_addr":shipper_addr[0],
                            "name":notify_name[0],
                            "notify_addr":notify_addr[0],
                            "port_load":port_load[0],
                            "port_dis":port_dis[0],
                            "marks":'Awaited',
                            "text_1":text_1[0],
                            "text_2":text_2[0],
                            "long_text_1":long_text_1[0],
                            "long_text_2":long_text_2[0],
                            "tot_rate":str(price[0]),
                            "blno": str(blno[0])

                        }
                doc.merge(**cust)
                doc.merge_rows('inv_desc', res)

                doc.write(f"./Invoice/{usr_input}_{file}.docx")
                pythoncom.CoInitialize()
                convert(f"./Invoice/{usr_input}_{file}.docx",f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/{usr_input}_{file}.pdf")
            except Exception as e:
                st.warning(e)
        if file_name is not None:
            for file in file_name:
                doc = DocxTemplate(file)
                doc.save(f"{file.name}")
                for x in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/"):
                    # st.write(x)
                    if x.endswith(".docx"):
                        doc = DocxTemplate(x)
                        x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/pck_lst?INVNO={}'.format(usr_input))
                        data = x.json()
                        for x in data['items']:
                            doc.render(x)
                            doc.save(
                                f"./Invoice/{usr_input}_{file.name}.docx")
                            pythoncom.CoInitialize()
                            convert(
                                f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/{usr_input}_{file.name}.docx",
                                f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/{usr_input}_{file.name}.pdf")
        st.success("Process completed.")
        zipf = zipfile.ZipFile(f'{usr_input}_doc.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/'):
            for file in files:
                if file[-4:] == 'docx':
                    zipf.write('./Invoice/' + file)
            zipf.close()
        zipf = zipfile.ZipFile(f'{usr_input}_pdf.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/'):
            for file in files:
                if file[-3:] == 'pdf':
                    zipf.write('./Invoice/' + file)
            zipf.close()
            def get_binary_file_downloader_html(bin_file, file_label='File'):
                with open(bin_file, 'rb') as f:
                    data = f.read()
                bin_str = base64.b64encode(data).decode()
                href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}"><button class="streamlit-button small-button primary-button ">{bin_file}</button></a>'
                return href
            pdf = st.markdown(get_binary_file_downloader_html(f'{usr_input}_pdf.zip', 'download'), unsafe_allow_html=True)
            doc = st.markdown(get_binary_file_downloader_html(f'{usr_input}_doc.zip', 'download'), unsafe_allow_html=True)
            if pdf and doc:
                for file in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/"):
                    os.remove(os.path.join(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/Invoice/{file}"))
            for file in os.listdir("C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/"):
                if file.endswith(".docx"):
                    os.remove(os.path.join(f"C:/Users/SoftChan.Arshad/Desktop/Garibsons Setup/Version 2.0/{file}"))
            btn = st.button("Clear")
            if btn:
                st.cache(func="clear")

