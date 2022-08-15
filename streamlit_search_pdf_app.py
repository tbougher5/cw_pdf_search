import os
import pandas as pd
import xlsxwriter
import fitz
import streamlit as st

#pdfDir = "D:\Dropbox\OBE\Python\Product Details"
fOut = 'Search_Products_List3.xlsx'
txtStr = 'WW-110'


#plFile = 'CurtainWall_PartsList.xlsx'
plFile = 'CW_Files.xlsx'
txtList = ['Horiztonal','Vertical','Mullion']
df = pd.read_excel(plFile)

pdfList = ["Reliance-TCIG-725-June2018.pdf"]
bolList = []
dtlList = []

#def convert_df(df):
#    return df.to_csv().encode('utf-8')


#csv = convert_df(df)



def run_search():

    row = 1
    ct = -1

    pdfDir = st.session_state.fldrPth
    fOut = st.session_state.fileOut
    txtStr = st.session_state.srchStr

    cols = ['Product Name','Full Name','Secondary Heading','Filename','Detail Heading 1','Detail Heading 2','Page Number','Search String']
    workbook = xlsxwriter.Workbook(fOut)
    worksheet = workbook.add_worksheet("Details Parts List") 
    worksheet.write_row(0,0,cols)

    dfOut = pd.DataFrame(columns = cols)
    dfLine = pd.DataFrame(columns = cols)

    for pdfFile in pdfList: #os.listdir("/"):
        ct += 1
        prd1 = ''
        ln1 = ''
        ln2 = ''
        
        if pdfFile.endswith('.pdf'):
            ff = False
            df2 = df.loc[(df['Filename'] == pdfFile)]
            pdfList.append(pdfFile)
            # print whole path of files
            pgNm = []
            pdNm = []
            pdf = fitz.open(pdfFile)
            pg = 0
            for page in pdf:
                mf = -1
                mf2 = -1
                pg += 1
                maxFnt = 0
                i = 0
                
                results = [] # list of tuples that store the information as (text, font size, font name) 
                dict = page.get_text("dict")
                blocks = dict["blocks"]
                for block in blocks:
                    if "lines" in block.keys():
                        spans = block['lines']
                        for span in spans:
                            data = span['spans']  
                            for lines in data:
                                #if keyword in lines['text'].lower(): # only store font information of a specific keyword
                                lineStr = lines['text'].replace('®','')
                                lineStr = lineStr.replace('™','')
                                
                                if pg == 1:
                                    for prodStr in df.columns:
                                        #Match the PDF with a product
                                        rawTxt = txtStr.split(' ')
                                        if len(rawTxt) == 1:
                                            bol1 = rawTxt[0].upper() in lineStr.upper()
                                        elif len(rawTxt) == 2: 
                                            bol1 = rawTxt[0].upper() in lineStr.upper() and rawTxt[1].upper() in lineStr.upper()
                                        elif len(rawTxt) == 3: 
                                            bol1 = rawTxt[0].upper() in lineStr.upper() and rawTxt[1].upper() in lineStr.upper() and rawTxt[2].upper() in lineStr.upper()
                                        elif len(rawTxt) == 4: 
                                            bol1 = rawTxt[0].upper() in lineStr.upper() and rawTxt[1].upper() in lineStr.upper() and rawTxt[2].upper() in lineStr.upper()  and rawTxt[3].upper() in lineStr.upper()
                                        if bol1:
                                            print('We found the product!')
                                            prd1 = prodStr
                                            ff = True
                                            #df2 = df[(df[prodStr].select_dtypes(include=['number']) != 0).any(1)]
                                            #df2 = df.loc[(df[prodStr] > 0)]
                                            break
                                if "BuildingEnvelope" in lines['text'] or 'Table of Contents' in lines['text']:
                                    be = True
                                elif '®' == lines['text']:
                                    be = False
                                elif '™' == lines['text']:
                                    be = False
                                else:
                                    results.append((lines['text'], lines['size'], lines['font']))                        
                                        
                                    if lines['size'] > maxFnt:
                                        maxFnt = lines['size']
                                        mf = i
                                        if pg == 1:
                                            ln1 = lineStr
                                        # lines['text'] -> string, lines['size'] -> font size, lines['font'] -> font name
                                    elif lines['size'] == maxFnt:
                                        mf2 = i
                                        if pg == 1:
                                            ln2 = lineStr 
                                    i += 1
                if pg > 2:              
                    #for prt in df2['OBE ITEM']:
                        for j in range(len(results)):
                            if txtStr in str(results[j]):  

                                dfOut = dfOut.append({'Product Name':df2['Product Name'].values[0], 'Full Name':df2['Full Name'].values[0], 'Secondary Heading':df2['Secondary Heading'].values[0],\
                                    'Filename': pdfFile, 'Detail Heading 1': results[mf][0], 'Detail Heading 2':results[mf2][0], 'Page Number':pg,\
                                        'Search String':txtStr, 'Count':1}, ignore_index=True)

                                row += 1               

            pdf.close()
            bolList.append(ff)
            if ct > 0:
                print('break')
                break
                            
    df3 = dfOut.groupby('Product Name').sum()#.drop(columns = 'Page Number')
    df4 = dfOut.groupby('Filename').sum()#.drop(columns = 'Page Number')

    with st.container():
        #col1, col2 = st.columns(2)
        col1edge, col1, col2edge = st.columns((1, 9, 1))
        col1.header('Full Search Results')
        col1.dataframe(dfOut)
        #col1.download_button("Download Results", csv, st.session_state.fileOut,"text/csv", key='browser-data')
        col1.header('Number of results by product')
        col1.dataframe(df3)
        col1.header('Number of results by file')
        col1.dataframe(df4)
# closing the pdf file object

# filePath is a string that contains the path to the pdf
ct = 0

st.set_page_config(layout="wide")

with st.container():
    col1edge, col1, col2edge = st.columns((1, 12, 1))
    col1.title('OBE Installation Manual Parts Search')
    #col1.caption('Calculate the carbon footprint of a fenestration system for commercial buildings across the US')
    col1.caption('')

with st.container():
    #col1, col2 = st.columns(2)
    col1edge, col1, col2, col3, col2edge = st.columns((1, 3, 3, 3, 1))
    col1.text_input("Folder Path", value = 'D:\Dropbox\OBE\Python\Product Details', key="fldrPth")
    col2.text_input("Search String", value = 'WW-110', key="srchStr")
    col3.text_input("Output File", value = 'pdf_search_results.csv', key="fileOut")
    col1.text('')
    col1.button(label='Calculate', key ='calc')
    #col1.button_calc = st.button(label='Calculate', key ='calc')


if st.session_state.calc:
    run_search()
    
