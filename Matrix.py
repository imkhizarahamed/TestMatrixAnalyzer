import pandas as pd
import plotly.express as px
import streamlit as st
from xlrd import *
from xlutils import *
from xlwt import *
from datetime import *
from PIL import Image




#nav_icon - page title - start
st.set_page_config(page_title="Test Matrix Analyzer", page_icon=":bar_chart:", layout="wide") #title
#nav_icon - page title - stop
#image - start
image = Image.open('logo.png')
col1, col2, col3,col4 = st.columns(4)
col2.image(image, width=620)
#image - stop
#title - start
st.markdown("<h1 style='text-align: center; color: Black;'>Test Matrix Analyzer</h1>", unsafe_allow_html=True)
#title - stop
#uploading file - start
uploaded_file = st.file_uploader("Upload Test Matrix File", type="xlsx")
if uploaded_file is not None:
#uploading file - stop
    dataframe = pd.read_excel(
        uploaded_file,
        sheet_name="Execution-ER"
    )
    df = pd.DataFrame(dataframe, columns=['P/F','Assigned To','Sl.No','Features','Date Executed'])
    st.header('Execution Status', anchor=None)
    c1,c2=st.columns(2)
    total=df['Sl.No'].count()
    actual=df['P/F'].count()
    percent=(actual/total)*100
    wr=str(round(percent,2))+' %'
    with c1:
        pf_values=df['P/F'].value_counts()#counting pass_fail casses
        if len(pf_values)==4:
            br_chart1_1=px.bar(pf_values,text_auto=True,color=['green','red','orange','darkblue'],
                        color_discrete_map="identity")
            br_chart1_1.update_layout(title_text='<b>Test Cases Pass/Failed<b>',title_x=0.5)
            br_chart1_1.update_layout(yaxis_title=None,xaxis_title=None)
            st.plotly_chart(br_chart1_1)
        elif len(pf_values)==3:  
            br_chart1=px.bar(pf_values,text_auto=True,color=['green','red','darkblue'],
                        color_discrete_map="identity")
            br_chart1.update_layout(title_text='<b>Test Cases Pass/Failed<b>',title_x=0.5)
            br_chart1.update_layout(yaxis_title=None,xaxis_title=None)
            st.plotly_chart(br_chart1)  
    
    with c2:
        dataframe_c = pd.read_excel(
        uploaded_file,
        sheet_name="Execution-ER"
    )
        dfx = dataframe_c.rename(columns={'Sl.No': 'Total Cases', 'P/F': 'Completed Cases'})
        completed_cases=pd.DataFrame(dfx, columns=['Total Cases','Completed Cases'])
        cc=completed_cases.count() 
        br_chart2=px.bar(cc,text_auto=True,title='Completed Cases',color=["red", "green"], color_discrete_map="identity")
        br_chart2.update_layout(title_text='<b>Total Cases/Completed Cases<b>',title_x=0.5)
        br_chart2.update_layout(yaxis_title=None,xaxis_title=None)
        br_chart2.update_traces(width=0.5)
        st.plotly_chart(br_chart2)  
    st.subheader('Percentage of Cases Completed till date: '+str(wr))
    
    dataframe_b= pd.read_excel(
        uploaded_file,
        sheet_name="Bugs Logged",   
    )
    #Bug Status--------------------------------------------------------------------------------------------------------------------
    st.header('Bugs Logged', anchor=None)
    c5,c6=st.columns(2)
    with c5:
        cr_bugs=dataframe_b['Build'].value_counts()
        br_chart3=px.bar(cr_bugs,text_auto=True,color=cr_bugs)
        br_chart3.update_layout(title_text='<b>Bugs logged in CR Builds<b>',title_x=0.5)
        br_chart3.update_coloraxes(showscale=False)
        br_chart3.update_layout(yaxis_title=None,xaxis_title=None)
        st.plotly_chart(br_chart3)
    with c6:
        comp_bugs=dataframe_b['Component'].value_counts()
        br_chart4=px.bar(comp_bugs,text_auto=True,color=comp_bugs)
        br_chart4.update_layout(title_text='<b>Bugs logged in Projects<b>',title_x=0.5)
        br_chart4.update_coloraxes(showscale=False)
        br_chart4.update_layout(yaxis_title=None,xaxis_title=None)
        st.plotly_chart(br_chart4)
    c7,c8=st.columns(2)
    with c7:
        bug_status=dataframe_b['Status'].value_counts()
        br_chart5=px.bar(bug_status,text_auto=True,color=bug_status)
        br_chart5.update_layout(title_text='<b>Bug Status<b>',title_x=0.5)
        br_chart5.update_coloraxes(showscale=False)
        br_chart5.update_layout(yaxis_title=None,xaxis_title=None)
        st.plotly_chart(br_chart5)  
    with c8:
        tester_bugs=dataframe_b['Reporter'].value_counts()
        br_chart8=px.bar(tester_bugs,text_auto=True,color=tester_bugs)
        br_chart8.update_layout(title_text='<b>Bugs Logged by testers<b>',title_x=0.5)
        br_chart8.update_coloraxes(showscale=False)
        br_chart8.update_layout(yaxis_title=None,xaxis_title=None)
        st.plotly_chart(br_chart8)
    #features covered------------------------------------------------------------------------------------------------
    st.header('Feature\'s Covered', anchor=None)
    c12,c13,c14,c15=st.columns(4)
    options = dataframe['Assigned To'].unique()
    filtered_df = dataframe[dataframe["Assigned To"].isin(options)]
    filter_df_fe=filtered_df['Features'].value_counts()
    br_chart9=px.bar(filter_df_fe,width=1500,text_auto=True,color=filter_df_fe)
    br_chart9.update_layout(title_text='<b>Feature\'s Covered till date<b>',title_x=0.5)
    br_chart9.update_coloraxes(showscale=False)
    br_chart9.update_layout(yaxis_title=None,xaxis_title=None)
    st.plotly_chart(br_chart9)
    st.subheader('Covered Features Table')
    st.write(filter_df_fe)
    #se;ect Options ----------------------------------------------------------------------------------------------
    st.header('Tester\'s Status', anchor=None)
    options = dataframe['Assigned To'].unique()
    multi_select_assigned=st.multiselect('Select Tester',options,default=options)  
    filtered_df = dataframe[dataframe["Assigned To"].isin(multi_select_assigned)]
    c10,c11=st.columns(2)
    with c10:
        filtered_df = dataframe[dataframe["Assigned To"].isin(multi_select_assigned)]
        filter_df_b=filtered_df['Assigned To'].value_counts()
        br_chart6=px.bar(filter_df_b,text_auto=True,color=filter_df_b)
        br_chart6.update_layout(title_text='<b>Test Cases Assigned<b>',title_x=0.5)
        br_chart6.update_coloraxes(showscale=False)
        br_chart6.update_layout(yaxis_title=None,xaxis_title=None)
        st.plotly_chart(br_chart6)
    with c11:
        filter_df=filtered_df['P/F'].value_counts()
        br_chart7=px.bar(filter_df,text_auto=True,color=filter_df)
        br_chart7.update_layout(title_text='<b>No.Of Test Cases Completed<b>',title_x=0.5)
        br_chart7.update_coloraxes(showscale=False)
        br_chart7.update_layout(yaxis_title=None,xaxis_title=None)
        st.plotly_chart(br_chart7)
    filtered_df = dataframe[dataframe["Assigned To"].isin(multi_select_assigned)]
    filter_df_f=filtered_df['Features'].value_counts()
    br_chart8=px.bar(filter_df_f,width=1500,text_auto=True,color=filter_df_f)
    br_chart8.update_coloraxes(showscale=False)
    br_chart8.update_layout(yaxis_title=None,xaxis_title=None)
    st.plotly_chart(br_chart8)
   
