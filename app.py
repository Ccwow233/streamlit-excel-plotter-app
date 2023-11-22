import streamlit as st
from streamlit_lottie import st_lottie
import requests
import openpyxl
import xlsxwriter
import time
import datetime
import os
import pandas as pd
import base64
from io import StringIO, BytesIO

def load_lottieurl(url):
    r = requests.get(url,verify=False)
    if r.status_code != 200:
        return None
    return r.json()

def progress_bar(seconds):
    progress_text = "Operation in progress. Please wait."
    progress_bar = st.progress(0,text=progress_text)
    for perc_completed in range(100):
        time.sleep(seconds)
        progress_bar.progress(perc_completed+1,text=progress_text)
    time.sleep(3)
    progress_bar.empty()
    st.success('Uploaded Successfully!')

@st.cache_data
def read_uploaded_file(upload_file):
    try:
        df = pd.read_excel(upload_file)
        return df
    except Exception as e:
        st.error(f"Error reading file:{e}")
        return None
                 
# Data cleaning for Posdata
def pos_cleaning():
    # divide six columns by 100 in posdata
    six_columns = ['BON_TPER', 'INCENT_TPER', 'OTHBON_TPER', 'BON_APER', 'INCENT_APER', 'OTHBON_APER']
    posdata_df[six_columns] = posdata_df[six_columns].div(100).where(posdata_df[six_columns].notna())
    # Calculate the tenure and categorise value into tenure distribution
    current_year = datetime.datetime.now().year
    posdata_df['Tenure'] = posdata_df['HIRE_YEAR'].apply(lambda x: current_year - x if pd.notna(x) else None)
    bins = [0, 1, 3, 6, 11, float('inf')]
    labels = [ '<1Y','1-2Y', '3-5Y', '6-10Y', '>10Y']
    posdata_df['Tenure Distribution'] = pd.cut(posdata_df['Tenure'], bins=bins, right = False, labels=labels, include_lowest=True)
    # Merge Posdata and Pname
    # add pname empty problem
    pname_keep = ['Job Code','Function', 'Sub-Function', 'Career Level by PC', 'CL-Sales','Career Level-Sales']
    merged_posdata = pd.merge(posdata_df, pname_df[pname_keep].fillna('NA'), left_on='TRS_POS_CODE', right_on='Job Code', how='left')
    compensation = ['CMP1','INCENT_TGT','BON_TGT','CMP3TGT_NEW','CMP3ACT','CMP5']
    basic_info = ['OBJECTID','EXCL_PML','SMI_CODE','ORGDATA_CPY_NAME','LOCATION_CN_MLS','YOUR_TITLE','TRS_POS_CODE','TRS_PNAME_POS_TITLE','TRS_POS_CLASS','Tenure','Tenure Distribution']
    merged_posdata = merged_posdata[basic_info + pname_keep + six_columns + compensation]
    # Filter the jobcodes in pname where comments is 'N'
    jobcodes_to_remove = pname_df.loc[pname_df['Comments'] == 'N', 'Job Code']
    # Remove the rows in posdata that match the jobcodes to remove
    posdata_new = merged_posdata[~merged_posdata['TRS_POS_CODE'].isin(jobcodes_to_remove)].copy()
    # Drop the 'Job Code' column from the merged_posdata dataframe
    posdata_new.drop('Job Code', axis=1,inplace=True)
    return posdata_new


def download_pos(posdata_cleaned):
    towrite = BytesIO()
    posdata_cleaned.to_excel(towrite,index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="POSDATA_NEW.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def download_cust(posdata_df):
    cust = BytesIO()
    posdata_df.to_excel(cust,index=False, header=True)
    cust.seek(0) 
    download3= st.download_button(
        label="Download Excel File",
        data=cust,
        file_name="Customise_groupby.xlsx",
        key='cust'
    )
    return download3

def download_posnew(posdata_cleaned):
    towrite = BytesIO()
    posdata_cleaned.to_excel(towrite,index=False, header=True)
    towrite.seek(0) 
    download2= st.download_button(
        label="Download Excel File",
        data=towrite,
        file_name="POSDATA_NEW.xlsx",
        key='download_button'
    )
    return download2
    
def groupby_rdpac(posdata_rdpac):
    # by function
    function_grp = posdata_rdpac.groupby(['SMI_CODE','Function'])['OBJECTID'].count().unstack()
    function_grp['Headcount'] = function_grp.sum(axis=1)
    # by sub-function
    subfunction_grp = posdata_rdpac.groupby(['SMI_CODE','Sub-Function'])['OBJECTID'].count().unstack()
    # by career level by pc
    clpc_grp = posdata_rdpac.groupby(['SMI_CODE','Career Level by PC'])['OBJECTID'].count().unstack()
    # by sub-function and cl-sales
    subsales_grp = posdata_rdpac.groupby(['SMI_CODE','Sub-Function','CL-Sales'])['OBJECTID'].count().unstack(level=[1]).unstack(level=[1])
    # by function and career level by pc
    funclpc_grp = posdata_rdpac.groupby(['SMI_CODE','Function','Career Level by PC'])['OBJECTID'].count().unstack(level=[1]).unstack(level=[1])
    # by smi_code and comp
    comp_grp = posdata_rdpac.groupby(['SMI_CODE']).agg(
        comp3act_sum=pd.NamedAgg(column='CMP3ACT', aggfunc='sum'),
        comp3tgt_mean=pd.NamedAgg(column='CMP3TGT_NEW', aggfunc='mean'),
        comp1_mean=pd.NamedAgg(column='CMP1', aggfunc='mean'),
        comp5_mean=pd.NamedAgg(column='CMP5', aggfunc='mean'),
        incentgt_sum=pd.NamedAgg(column='INCENT_TGT', aggfunc='sum')
    )
    # by function and incent_tgt
    funincent_grp = posdata_rdpac.groupby(['SMI_CODE','Function'])['INCENT_TGT'].mean().unstack()
    # by cl-sales and bon_tgt
    salesbon_grp = posdata_rdpac.groupby(['SMI_CODE','CL-Sales'])['BON_TGT'].mean().unstack()
    # by subfunction and cl-sales and incent_tgt
    subsalesincent_grp = posdata_rdpac.groupby(['SMI_CODE','Sub-Function','CL-Sales'])['INCENT_TGT'].mean().unstack(level=[1]).unstack(level=[1])
    # by function and career level by pc and bon_tper
    funsclpctper_grp = posdata_rdpac.groupby(['SMI_CODE','Function','Career Level by PC'])['BON_TPER'].mean().unstack(level=[1]).unstack(level=[1])
    # by subfunction and career level by pc and incent_tper
    subclpctper_grp = posdata_rdpac.groupby(['SMI_CODE','Sub-Function','Career Level by PC'])[['INCENT_TPER','BON_TPER']].mean().unstack(level=[1]).unstack(level=[1])
    return [function_grp,subfunction_grp,clpc_grp,subsales_grp,funclpc_grp,comp_grp,funincent_grp,salesbon_grp,subsalesincent_grp,funsclpctper_grp,subclpctper_grp]
    
def download_rdpac(rr_list,sheet_names,file_name='Grouped_Data.xlsx'):
    output = BytesIO()
    # Create an Excel writer object
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    # Export the DataFrames to separate worksheets
        for sheet_name,df in zip(sheet_names,rr_list):
            df.to_excel(writer,sheet_name=sheet_name)
    output.seek(0)
    download1= st.download_button(
        label="Download File",
        data=output,
        file_name=file_name,
        key='rdpac'
    )
    return download1

def headcount_cal():
    df_grouped = posdata_df.groupby(groupby_columns)['OBJECTID'].count().reset_index()
    return df_grouped

def comp_cal_per(ratio,posdata_df):
    posdata_df['Percentile']=posdata_df.groupby(groupby_columns)[output_columns].transform(lambda x:pd.qcut(x,q=ratio),)
    #df_grouped = posdata_df.groupby(by=[groupby_columns], as_index=False)[output_columns].sum()
    perc = posdata_df['Percentile']
    percentile_df = posdata_df[groupby_columns+output_columns+perc]
    return percentile_df

def comp_cal_sum(posdata_df):
    #df_grouped = posdata_df.groupby(by=[groupby_columns], as_index=False)[output_columns].sum()
    df_grouped = posdata_df.groupby(groupby_columns)[output_columns].sum().reset_index()
    return df_grouped

def comp_cal_mean(posdata_df):
    #df_grouped = posdata_df.groupby(by=[groupby_columns], as_index=False)[output_columns].sum()
    df_grouped = posdata_df.groupby(groupby_columns)[output_columns].mean().reset_index()
    return df_grouped

st.set_page_config(page_title='HCE Data Explorer',page_icon=':partying_face',layout='wide')

lottie_url = load_lottieurl('https://lottie.host/2d414cda-2d5d-4823-a3aa-a25107a86605/bFn1VfLAUS.json')

with st.sidebar:
    st.header('Tutorial')
    st.write('##')
    st.markdown('__Step 1__ Provide the filepath,then match the right file from dropdown list in the selectbox')
    st.markdown('__Step 2__ Posdata Data Cleaning and merge with Pname,then you could choose to download new posdata file')
    st.markdown('__Step 3__ Generate Raw Data and then you could choose to download the grouped data')
                
with st.container():	
    left_column,right_column = st.columns(2)	
    with left_column:	
        st.subheader("Hello 👋")	
        st.title('Weclcome to :blue[HCE Data Explorer]')	
    with right_column:	
        st.lottie(lottie_url,height=300)	

with st.container():
    st.write('---')
    st.header('Step 1: _Provide Data Source_')

with st.container():
    file1 = st.file_uploader('Choose POSDATA file',type='xlsx',key='1')
    if file1 is not None:
        posdata_df = read_uploaded_file(file1)
        st.success('Upload it Successfully!')
    file2 = st.file_uploader('Choose PNAME file',type='xlsx',key='2')
    if file2 is not None:
        pname_df = read_uploaded_file(file2)
        st.success('Upload it Successfully!')
    st.markdown('🙄 **If you wanna customise groupby,please go to step 3**')

with st.container():
    st.write('---')
    st.write('##')
    left_column1, right_column1 = st.columns(2)
    with left_column1:
        st.header('Step 2.1: _Data Cleaning (Only for RDPAC)_')
        st.markdown('🙄 **Posdata Cleaning Button**')
        if st.button('Click me',key='cleaning'):
            pos = pos_cleaning()
            st.write('Complete it!')       
            st.subheader('Downloads:')
            download_posnew(pos)
    with right_column1:
        st.header('Step 2.2: _Generate Raw Data (Only for RDPAC)_')
        st.markdown('😋 **Default RDPAC Raw Data**')
        if st.button('Run me',key='Generating'):
            rr = pos_cleaning()
            rr_list = groupby_rdpac(rr)
            st.write('Complete it!') 
            sheet_names = ['By Function','By Subfunction','By Career Level by PC',
                'By Subfunction plus Sales','By Function plus CL','COMP',
                'By Function plus Incent','By cl plus Bon','By SubFunction cl plus Incent',
                'By Function PC plus Bon','By SubFunction PC plus Incent']     
            st.subheader('Downloads:')
            download_rdpac(rr_list,sheet_names)

with st.container():
    st.write('---')
    st.header('Step 3: _Customised Group By_')   
    st.markdown('**Choose what you want !!!**')
    groupby_columns= st.multiselect(
    'What are your groupby combination? (Multiselect from your posdata)',
    posdata_df.columns.tolist())
    st.write('You selected:', groupby_columns)
    genre = st.radio(
"Which one do you want to caculate?",
['Headcount','Compensation']
)   
    if genre == 'Headcount':
        hc_df = headcount_cal()
        st.subheader('Downloads:')
        download_cust(hc_df)
    else:    
        output_columns = st.multiselect(
"What compensation do you want to caculate?",
['CMP1','CMP2','CMP3ACT','CMP3TGT_NEW',
 'CMP4','CMP5_NOLTI','CMP5','CMP5TGT_NOLTI'])
        exclude_flag = st.selectbox('Choose your survey_exclude_flag:',posdata_df.columns.tolist())
        method = st.radio(
    "How do you caculate?",
    ['Sum','mean']
    )  
        #if method == 'Percentile':
           # ratio = st.number_input("Insert a number", value=None, placeholder="Type a number...")
            #cmp_per = comp_cal_per(ratio,posdata_df)
            #st.subheader('Downloads:')
            #download_cust(cmp_per)###
        if method == 'Sum':
            if exclude_flag is not None:
                posdata_exl = posdata_df[posdata_df[exclude_flag].isnull()]
                comp_cal = comp_cal_sum(posdata_exl)
                st.subheader('Downloads:')
                download_cust(comp_cal)
            else:
                comp_cal = comp_cal_sum()
                st.subheader('Downloads:')
                download_cust(comp_cal)
        else:
            if exclude_flag is not None:
                posdata_exl = posdata_df[posdata_df[exclude_flag].isnull()]              
                comp_cal = comp_cal_mean(posdata_exl)
                st.subheader('Downloads:')
                download_cust(comp_cal)
            else:
                comp_cal = comp_cal_mean()
                st.subheader('Downloads:')
                download_cust(comp_cal)






    






            
