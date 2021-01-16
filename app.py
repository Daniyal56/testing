# Using python 3.6.1
# importing libraries
import streamlit as st
import os,base64,datetime
import pandas as pd
import numpy as np
from tabula import convert_into,read_pdf
from streamlit.components.v1 import html
from convert_to_pdf import file_converter
from load import local_css
import datetime
from datetime import timedelta
date = datetime.datetime.now()



# st.set_page_config(page_title='Demo',layout='wide',initial_sidebar_state='auto')
local_css("style.css")

# Sidebar setup
def app():


    # t = "<div><span class='highlight blue' style='text-align:center; font-family:Roboto; font-size:20.5px;'><b>BANK RECONCILIATION</b> </span></div>"
    # st.sidebar.markdown(t, unsafe_allow_html=True)
    #
    #
    # st.sidebar.text('')
    # st.sidebar.text('')
    cont = st.sidebar.beta_expander('Type Section')
    lov_selection = ['Comparison','Converter - PDF/CSV']
    lov_selection_col = cont.selectbox('Please Choose',lov_selection)
    t = f"<p style='color:black;'><b>You selected</b> <span class='highlight red' style='font-family:Roboto;'>{lov_selection_col}</span></p>"
    # st.sidebar.markdown(t, unsafe_allow_html=True)
    cont.markdown(t,unsafe_allow_html=True)
    st.sidebar.text('')
    st.sidebar.text('')
    # function to download CSV

    def get_binary_file_downloader_html(bin_file, file_label='File'):
        with open(bin_file, 'rb') as f:
            data = f.read()
        bin_str = base64.b64encode(data).decode()
        href = f'<a style="font-family:Roboto;" href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">{bin_file}</a>'
        return href

    # Main Page Setup
    #Set Lov
    if lov_selection_col == 'Comparison':
        t = "<h1 style='text-align:center;'><span class='highlight blue' style='text-align:center; font-family:Roboto; font-size:20.5px;'><b><u>COMPARISON CSV FILE's</u></b> </span></h2>"
        st.markdown(t, unsafe_allow_html=True)

        # st.markdown(navbar,unsafe_allow_html=True)
        st.write(' ')
        st.write(' ')
        # bank_note = '<p style="font-family:Roboto;"><b>Please select your Company file first.</b></p>'
        # st.sidebar.markdown(bank_note, unsafe_allow_html=True)
        cont_sec = st.sidebar.beta_expander('Uploading Section')
        try:
            files = cont_sec.file_uploader('Please upload files', accept_multiple_files=True, type='csv')

        except IndexError as e:
            st.warning('Please upload maximum two files\nIf you are still getting an error then please check you file extension')
        # Rendering files name page


        # st.markdown('<p style="font-family:Roboto;"> <u>You uploaded <b>{}</b> & <b>{}</b></u></p>'.format(files[0].name,files[1].name),unsafe_allow_html=True)
        t = f"<div><b>You uploaded</b> <span class='highlight blue'><b>{files[0].name}</b></span> </span> <b>&</b> <span class='highlight red'><b>{files[1].name}</b></span></div>"

        st.markdown(t, unsafe_allow_html=True)

        ## Reading CSV's
        st.write('')
        company = pd.read_csv(files[0])
        bank = pd.read_csv(files[1])
        dir = 'C:/Users/Daniyal/Desktop/Bank Reconcilation/backup'
        x = np.random.randint(1,10000000,1)
        company.to_csv(f'{dir}/{x}_{files[0].name}')
        bank.to_csv(f'{dir}/{x}_{files[1].name}')
        # Bank Columns
        st.write('')
        st.write('')
        bnk_chq_no,bnk_debit,bnk_crdt = st.beta_columns(3)
        with bnk_chq_no:
            bnk_chq_no_col = st.selectbox('Please type CHQ Column of {}'.format(files[1].name),options=list(bank.columns))
        with bnk_debit:
            bnk_debit_col = st.selectbox('Please type Debit Column of {}'.format(files[1].name),options=list(bank.columns))
        with bnk_crdt:
            bnk_crdt_col = st.selectbox('Please type Credit Column of {}'.format(files[1].name),options=list(bank.columns))
        # Company Columns
        cmp_chq_no, cmp_debit, cmp_crdt = st.beta_columns(3)
        with cmp_chq_no:
            cmp_chq_no_col = st.selectbox('Please type CHQ Column of {}'.format(files[0].name),options=list(company.columns))
        with cmp_debit:
            cmp_debit_col = st.selectbox('Please type Debit Column of {}'.format(files[0].name),options=list(company.columns))
        with cmp_crdt:
            cmp_crdt_col = st.selectbox('Please type Credit Column of {}'.format(files[0].name),options=list(company.columns))

        #Bank difference
        # t = "<h1 style='text-align:center;'><span class='highlight red' style='text-align:center; font-family:Roboto; font-size:30.5px;'><b><u>Record\'s Section</u>:</b> </span></h1>"
        # st.markdown(t, unsafe_allow_html=True)
        if st.checkbox('Submit',key='s'):
            indexes = bank[~(bank[bnk_chq_no_col].isin(company[cmp_chq_no_col]) &
                bank[bnk_crdt_col].isin(company[cmp_debit_col]) | bank[bnk_debit_col].isin(company[cmp_crdt_col]) & bank[bnk_chq_no_col].isin(company[cmp_chq_no_col]))]
            length = len(indexes)
            st.markdown(f'<h3 style="font-family:Roboto;"><b><u>Unmatched Records in {files[1].name}</u> => (Record\'s found {str(length)})</b></h3>',
                        unsafe_allow_html=True)
            # t = f"<div><b style='font-family:Roboto; font-size:30;'>Unmatched Records in</b> <span class='highlight blue'><b>{files[1].name}</b> </span>&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;<b><u>Record\'s found</u></b> <span class='highlight blue'> <b>{str(length)}</b></span></div>"
            # st.markdown(t,unsafe_allow_html=True)

            # st.write('--'*50)
            st.markdown('<hr>',unsafe_allow_html=True)
            indexes['status'] = ''
            header = st.beta_columns(4)
            status = []
            lov = ['--Select--', 'NSF', 'Dishonoured', 'Outstanding','Unpresented']
            header[0].markdown('<h4 style="text-align:center; font-family:Roboto;"><b><u>CHQ</u></b><h4>', unsafe_allow_html=True)
            header[1].markdown('<h4 style="text-align:center; font-family:Roboto;"><b><u>Debit</u></b><h4>', unsafe_allow_html=True)
            header[2].markdown('<h4 style="text-align:center; font-family:Roboto;"><b><u>Credit</u></b><h4>', unsafe_allow_html=True)
            header[3].markdown('<h4 style="text-align:center; font-family:Roboto;"><b><u>Status</u></b><h4>', unsafe_allow_html=True)
            for x, y in enumerate(indexes[bnk_chq_no_col]):
                header[0].text_input(f'', value=y, key=f'{x}')
            for idx, deb in enumerate(indexes[bnk_debit_col]):
                header[1].text_input(f'', value=deb, key=f'{idx}')
            for idx_, cre in enumerate(indexes[bnk_crdt_col]):
                header[2].text_input(f'', value=cre, key=f'{idx_}_')
            for idx__, s in enumerate(indexes['status']):
                x = header[3].selectbox(f'', lov, key=f'_{idx__}')
                status.append(x)
            indexes['status'] = status

            indexes.to_csv(f'Unmatched_{files[0].name}',index=False)
            sub_btn, v, t, y, u, i, o, sb_btn = st.beta_columns(8)
            with sb_btn:
                sb__btn = st.button("Convert-PDF")
            indexes = company[~(company[cmp_chq_no_col].isin(bank[bnk_chq_no_col]) &
                company[cmp_debit_col].isin(bank[bnk_crdt_col])) | ~(company[cmp_crdt_col].isin(bank[bnk_debit_col]) & company[cmp_chq_no_col].isin(bank[bnk_chq_no_col]))]
            length = len(indexes)
            bnk_diff_expndr = st.beta_expander('Unmatched Records in {} '.format(files[0].name))
            t = f"<div>&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;<b><u>Record\'s found</u></b> <span class='highlight blue'> <b>{str(length)}</b></span></div>"
            bnk_diff_expndr.markdown(t, unsafe_allow_html=True)
            bnk_diff_expndr.table(indexes)

            indexes.to_csv(f'Unmatched_{files[1].name}',index=False)

            #Matched Records
            indexes = bank[(bank[bnk_chq_no_col].isin(company[cmp_chq_no_col]) & bank[bnk_crdt_col].isin(company[cmp_debit_col])) | (bank[bnk_debit_col].isin(company[cmp_crdt_col]) & bank[bnk_chq_no_col].isin(company[cmp_chq_no_col]))]
            length = len(indexes)
            mtch_expndr = st.beta_expander(
                'Matched Records')
            t = f"<div>&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;<b><u>Record\'s found</u></b> <span class='highlight blue'> <b>{str(length)}</b></span></div>"
            mtch_expndr.markdown(t, unsafe_allow_html=True)
            mtch_expndr.table(indexes)
            indexes.to_csv('Matched_Records.csv',index=False) #Converting into CSV

            st.text('')
            st.text('')
            t = "<p style='text-align:center;'><span class='highlight red' style='text-align:center; font-family:Roboto; font-size:20.5px;'><b><u>Download Section</u>:</b> </span></p>"
            st.markdown(t, unsafe_allow_html=True)
            cmp_diff,bnk_diff,matched = st.beta_columns(3)
            downld_sec = st.beta_expander('Download CSV Files!')
            with cmp_diff:
                downld_sec.markdown(get_binary_file_downloader_html(f'Unmatched_{files[0].name}', ''), unsafe_allow_html=True)
            with bnk_diff:
                downld_sec.markdown(get_binary_file_downloader_html(f'Unmatched_{files[1].name}', ''), unsafe_allow_html=True)
            with matched:
                downld_sec.markdown(get_binary_file_downloader_html('Matched_Records.csv', ''), unsafe_allow_html=True)
                    # Working for PDF
            downld_pdf = st.beta_expander('Download PDF Files!')
            cmp_diff_pdf, bnk_diff_pdf, matched_pdf = st.beta_columns(3)


            if sb__btn:
                file_converter(csv_file=f'Unmatched_{files[0].name}')
                with cmp_diff_pdf:
                    d = downld_pdf.markdown(get_binary_file_downloader_html(f'./pdf_folder/Unmatched_{files[0].name}_.pdf', ''), unsafe_allow_html=True)
                file_converter(csv_file=f'Unmatched_{files[1].name}')
                with bnk_diff_pdf:
                    downld_pdf.markdown(get_binary_file_downloader_html(f'./pdf_folder/Unmatched_{files[1].name}_.pdf', ''), unsafe_allow_html=True)
                file_converter(csv_file='Matched_Records.csv')
                with matched_pdf:
                    downld_pdf.markdown(get_binary_file_downloader_html('./pdf_folder/Matched_Records.csv_.pdf', ''), unsafe_allow_html=True)

    elif lov_selection_col == 'Converter - PDF/CSV':
        try:
            # header = '<h1 style="text-align:center; font-family:Roboto; font-size:20.5px;"><b><u>PDF</u> <u>CONVERTER</u></b></h1>'
            # st.markdown(header,unsafe_allow_html=True)
            t = "<h1 style='text-align:center;'><span class='highlight blue' style='text-align:center; font-family:Roboto; font-size:20.5px;'><b><u>PDF</u> <u>CONVERTER</u></b> </span></h2>"
            st.markdown(t, unsafe_allow_html=True)
            file = st.file_uploader('Please upload your pdf file',type='pdf')
            para_cont = '<p style="font-family:Roboto;">You uploaded the file <span class="highlight blue" style="text-align:center; font-family:Roboto; font-size:15.5px;"><b>{}</b></span>, Please give us a second\'s to track the tables in your pdf.</p>'.format(file.name)
            st.markdown(para_cont,unsafe_allow_html=True)
            st.text('')
            st.text('')
            #writing code for pdf
            dwnload,pg,viw= st.beta_columns(3)
            if file:
                pdf_viewer,page,download = st.beta_columns(3)
                pdf_area = st.beta_expander("Download Converted CSV.")
                convert_into(file, f'{file.name}.csv', stream=True, pages='all')
                pdf_area.markdown(get_binary_file_downloader_html(f'{file.name}.csv'), unsafe_allow_html=True)

        except Exception as e:
            pass