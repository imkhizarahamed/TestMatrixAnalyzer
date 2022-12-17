import pandas as pd
import plotly.express as px
import streamlit as st
from xlrd import *
from xlutils import *
from xlwt import *
from datetime import *
from PIL import Image
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib,ssl
from fpdf import FPDF
import base64
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import dash_core_components as dcc
import plotly.subplots as sp



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
    
    #-----------------------------------------------------------------------------

    with open('config.yaml') as file:
        config = yaml.load(file, Loader=SafeLoader)

# Creating the authenticator object
    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'], 
        config['cookie']['key'], 
        config['cookie']['expiry_days'],
        config['preauthorized']
    )

    #hashed_passwords = stauth.Hasher(['123']).generate()


    # creating a login widget
    name, authentication_status, username = authenticator.login('Please Login to download/send report', 'main')
    if authentication_status:
            #start-----------------------------------------------------------------------------------------------------------
            df = pd.DataFrame(dataframe, columns=['P/F','Assigned To','Sl.No','Features','Date Executed'])
            st.header('Download PDF Report: ', anchor=None)
            c1,c2=st.columns(2)
            total=df['Sl.No'].count()
            actual=df['P/F'].count()
            percent=(actual/total)*100
            wr=str(round(percent,2))+' %'
            pf_values=df['P/F'].value_counts()#counting pass_fail casses
            dataframe_c = pd.read_excel(
                uploaded_file,
                sheet_name="Execution-ER"
            )
            dfx = dataframe_c.rename(columns={'Sl.No': 'Total Cases', 'P/F': 'Completed Cases'})
            completed_cases=pd.DataFrame(dfx, columns=['Total Cases','Completed Cases'])
            cc=completed_cases.count() 
            #fig1
            pf_values=df['P/F'].value_counts()
            figure1=px.bar(pf_values,text_auto=True,color=pf_values)
            figure1.update_layout(title_text='<b>Test Cases Pass/Failed<b>',title_x=0.5)
            figure1.update_layout(yaxis_title=None,xaxis_title=None) 
            #fig2
            figure2 = px.bar(cc,text_auto=True,color=cc)
            figure2.update_coloraxes(showscale=False)
            figure2.update_layout(yaxis_title=None,xaxis_title=None) 
            #fig3
            dataframe_b= pd.read_excel(
                uploaded_file,
                sheet_name="Bugs Logged",   
            )
            cr_bugs=dataframe_b['Build'].value_counts()
            figure3=px.bar(cr_bugs,text_auto=True,color=cr_bugs)
            figure3.update_layout(title_text='<b>Bugs logged in CR Builds<b>',title_x=0.5)
            figure3.update_coloraxes(showscale=False)
            figure3.update_layout(yaxis_title=None,xaxis_title=None)
            #fig4
            comp_bugs=dataframe_b['Component'].value_counts()
            figure4=px.bar(comp_bugs,text_auto=True,color=comp_bugs)
            figure4.update_layout(title_text='<b>Bugs logged in Projects<b>',title_x=0.5)
            figure4.update_coloraxes(showscale=False)
            figure4.update_layout(yaxis_title=None,xaxis_title=None)
            #fig5
            bug_status=dataframe_b['Status'].value_counts()
            figure5=px.bar(bug_status,text_auto=True,color=bug_status)
            figure5.update_layout(title_text='<b>Bug Status<b>',title_x=0.5)
            figure5.update_coloraxes(showscale=False)
            figure5.update_layout(yaxis_title=None,xaxis_title=None)
            #fig6
            tester_bugs=dataframe_b['Reporter'].value_counts()
            figure6=px.bar(tester_bugs,text_auto=True,color=tester_bugs)
            figure6.update_layout(title_text='<b>Bugs Logged by testers<b>',title_x=0.5)
            figure6.update_coloraxes(showscale=False)
            figure6.update_layout(yaxis_title=None,xaxis_title=None)
            #traces
            figure1_traces = []
            figure2_traces = []
            figure3_traces = []
            figure4_traces = []
            figure5_traces = []
            figure6_traces = []
            for trace in range(len(figure1["data"])):
                figure1_traces.append(figure1["data"][trace])

            for trace in range(len(figure2["data"])):
                figure2_traces.append(figure2["data"][trace])
            
            for trace in range(len(figure3["data"])):
                figure3_traces.append(figure3["data"][trace])
            
            for trace in range(len(figure4["data"])):
                figure4_traces.append(figure4["data"][trace])
            
            for trace in range(len(figure5["data"])):
                figure5_traces.append(figure5["data"][trace])
            
            for trace in range(len(figure6["data"])):
                figure6_traces.append(figure6["data"][trace])
            #Create a 1x2 subplot
            this_figure = make_subplots(rows=3, cols=2,subplot_titles=('Test Cases Pass/Failed',
                                            'Total Completed Cases',
                                            'Bugs logged in CR Builds',
                                            'Bugs logged in Projects',
                                            'Bug Status',
                                            'Bugs Logged by Testers')) 
            
            # Get the Express fig broken down as traces and add the traces to the proper plot within in the subplot
            for traces in figure1_traces:
                this_figure.append_trace(traces, row=1, col=1)
            for traces in figure2_traces:
                this_figure.append_trace(traces, row=1, col=2)
            for traces in figure3_traces:
                this_figure.append_trace(traces, row=2, col=1)    
            for traces in figure4_traces:
                this_figure.append_trace(traces, row=2, col=2)  
            for traces in figure5_traces:
                this_figure.append_trace(traces, row=3, col=1) 
            for traces in figure6_traces:
                this_figure.append_trace(traces, row=3, col=2) 
            #names = {'Plot 1':'2016', 'Plot 2':'2017', 'Plot 3':'2018', 'Plot 4':'2019'}
            #this_figure.for_each_annotation(lambda a: a.update(text = a.text + ': ' + names[a.text]))
            #the subplot as shown in the above image
            final_graph = dcc.Graph(figure=this_figure)
            
            st.plotly_chart(this_figure)
            # Load the data
            @st.experimental_memo
            def load_data():
                    return pd.DataFrame(dataframe)

            # Create and cache a Plotly figure
            @st.experimental_memo
            def create_figure(df):
                fig=this_figure
                fig.update_layout(title_text="Execution Report")
                fig.update_coloraxes(showscale=False)
                return fig
            df = load_data()
            fig = create_figure(df)
            # Create an in-memory buffer
            buffer = io.BytesIO()
            layout = go.Layout(
            autosize=True
            )   

            # Save the figure as a pdf to the buffer
            fig.write_image(file=buffer, format="pdf",width=1980, height=2080)
            

            # Download the pdf from the buffer
            st.download_button(
                label="Download PDF",
                data=buffer,
                file_name="figure.pdf",
                mime="application/pdf",
            )

 
            
        #################################################################################### 
            
            st.title("Send Mail")
            #uploaded_file = st.file_uploader("Choose a file")
            temp_file = st.file_uploader("Enter file here!")
            if temp_file: 
                temp_file_contents = temp_file.read()

            #if st.button("Save as working file"):   
            #   with open("ON_DISK_FILE.extension","wb") as file_handle:
            #      file_handle.write(temp_file_contents)


            result= st.button('Click To Send Mail')
            st.write(result)
            if result:

                my_email= "ShaikMohammad.Khizar@contractor.tranetechnologies.com"
                password= "King@6775"
                from_addr='ShaikMohammad.Khizar@contractor.tranetechnologies.com'
                to_addrs=st.text_input('Enter Sender Address')
                msg=st.text_input('Enter Message')
                server = smtplib.SMTP('smtp-mail.outlook.com', 587)
                server.starttls()
                server.ehlo()
                server.login(my_email, password)
                server.sendmail(from_addr,to_addrs,msg)
                server.close()
            authenticator.logout('Logout', 'main')
    elif authentication_status == False:
            st.error('Username/password is incorrect')
    
    
    ########################
    
    