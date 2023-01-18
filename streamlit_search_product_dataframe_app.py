import pandas as pd
import streamlit as st
import pickle

prdDFlist = [['All Products','DF_All_Products_v4a.sav'], ['Architectural Windows', 'DF_Architectural_Windows_v4.sav'], ['Curtain Wall', 'DF_Curtain_Wall_v4.sav']\
             ,['Entrances','DF_Entrances_v4.sav'],['Storefront','DF_Storefront_v4.sav'],['Window Wall','DF_Window_Wall_v4.sav'],['Sun Controls', 'DF_Sun_Controls_v4.sav']]
prdCat = ['All Products', 'Architectural Windows', 'Curtain Wall','Entrances','Storefront','Window Wall','Sun Controls']
prd = 'Curtain Wall'
prdNum = 2

bolList = []
dtlList = []

def convert_df(df):
    
    return df.to_csv().encode('utf-8')

def run_search():

    row = 1
    ct = -1

    #pdfDir = st.session_state.fldrPth
    fOut = st.session_state.fileOut
    
    prdCat = st.session_state.pc
    
    #if st.session_state.exStr == True:
      #txtStr = st.session_state.srchStr + '\b'
    #else:
    txtStr = st.session_state.srchStr
    
    if st.session_state.hypStr == True:
      txtStr = txtStr.replace('-','*')

    for i in range(len(prdDFlist)):
        if prdCat in prdDFlist[i][0]:
            dfFile = prdDFlist[i][1]
            break

    #cols = ['Product Name','Full Name','Secondary Heading','Filename','Detail Heading 1','Detail Heading 2','Page Number','Search String']
    #workbook = xlsxwriter.Workbook(fOut)
    #worksheet = workbook.add_worksheet("Details Parts List") 
    #worksheet.write_row(0,0,cols)
    dfPrd = pickle.load(open(dfFile,'rb'))
    #print(dfPrd)
    #print(prd)
    if prdCat == 'All Products':
      if st.session_state.exStr == True:
        dfOut = dfPrd.loc[(dfPrd['Text'].str.contains(txtStr+'\b',case=False, na=False, regex=True))]
      else:
        dfOut = dfPrd.loc[(dfPrd['Text'].str.contains(txtStr,case=False, na=False, regex=True))]
    else:
        dfOut = dfPrd.loc[(dfPrd['Text'].str.contains(txtStr,case=False, na=False, regex=True)) & (dfPrd['Product Category'] == prdCat)]

    dfOut['Count'] = dfOut['Count'].astype('int')
    df3 = dfOut.groupby('Product').sum().drop(columns = 'Page Number')
    df4 = dfOut.groupby('Filename').sum().drop(columns = 'Page Number')
    
    #dfLine = pd.DataFrame(columns = cols)
    #dfOut.to_csv(fOut)
    
    csv = convert_df(dfOut)

    with st.container():
        #col1, col2 = st.columns(2)
        col1edge, col1, col2edge = st.columns((1, 9, 1))
        col1.header('Full Search Results')
        col1.download_button("Download Results", csv, st.session_state.fileOut,"text/csv", key='browser-data')
        col1.dataframe(dfOut)
        col1.header('Number of results by product')
        col1.dataframe(df3)
        col1.header('Number of results by file')
        col1.dataframe(df4)
        col1.text(txtStr)
# closing the pdf file object

# filePath is a string that contains the path to the pdf
ct = 0

st.set_page_config(layout="wide")

with st.container():
    col1edge, col1, col2edge = st.columns((1, 12, 1))
    col1.title('Web Details Parts Search')
    col1.caption('')
    col1.caption('Contains web details, installation manuals, and Florida Product Approvals')
    col1.caption('Currently some details missing for Entrances and Windows due to PDF issues')
    col1.caption('')

with st.container():
    #col1, col2 = st.columns(2)
    col1edge, col1, col2, col3, col2edge = st.columns((1, 3, 3, 3, 1))
    col1.selectbox('Product Category', prdCat, index=0, key='pc', help=None, on_change=None, args=None, kwargs=None, disabled=False)
    col2.text_input("Search String", value = 'WW-110', key="srchStr")
    col3.text_input("Output File", value = 'pdf_search_results.csv', key="fileOut")
    col1.text('')
    col2.text('')
    col3.text('')
    col1.checkbox('Exact part only (e.g., GP-100 but not GP-1001)', key='exStr', value=True)
    col2.checkbox('Include results without hyphen or space', key='hypStr', value=True)
    col1.text('')
    col1.button(label='Calculate', key ='calc')
    #col1.button_calc = st.button(label='Calculate', key ='calc')

if st.session_state.calc:
    run_search()
    
