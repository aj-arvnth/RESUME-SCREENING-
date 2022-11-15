import streamlit as st
from deta import Deta
from textblob import TextBlob
import pandas as pd
import smtplib as s
import os
from win32com.client import Dispatch

finalscore=[]
def speak(str):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str)

def final(id, finallist):
    st.write(id)
    usermail="tidelhr0305@gmail.com"
    password="rmmgguvjupragcaw"
    recipient=finallist[id]["mail"]
    subject="Call For Interview!!"
    message="Congrats!!You have been shortlisted for the interview. Date will be announced later"
    
    
    try:
        connection=s.SMTP('smtp.gmail.com',587)
        connection.starttls()
        connection.login(usermail,password)
        fullmsg="Subject:{}\n\n{}".format(subject,message)
        connection.sendmail(usermail,recipient,fullmsg)
        connection.quit()
        st.success("Email Sent Successfully")
        speak("Confirmation mail is successfully sent to your shortlisted candidate")
        
        
    except Exception as e:
    
        st.write(e)
def commonfunction(finallist):
    i=0
    while(i<len(finallist)):
        
        st.markdown("<div id='box' style='margin-bottom: 20px; padding: 20px; margin-top: 40px; height: 250px; width: 500px; border: 2px solid white; border-radius:12px; padding-left: 30px;'><h5><u style='color:#5eb5a8;'>Name:</u>&nbsp;&nbsp;&nbsp;&nbsp;"+finallist[i]["Name"]+"<br><u style='color:#5eb5a8;'>Skills:</u>&nbsp;&nbsp;&nbsp;&nbsp;"+finallist[i]["skills"]+"<br><u style='color:#5eb5a8;'>About him:</u>&nbsp;&nbsp;&nbsp;&nbsp;"+finallist[i]["aboutme"]+"<br><u style='color:#5eb5a8;'>Resume Score:</u>&nbsp;&nbsp;&nbsp;&nbsp;"+str(finalscore[i])+" </h5></div>",unsafe_allow_html=True)
        if(st.button("Send Confirmation Mail to Applicant "+str(i+1))):
            final(i,finallist)
        i=i+1
    
  
menu=["Student","Admin"]
choice=st.sidebar.selectbox("Menu",menu)
deta = Deta("d0l89iaj_HxpizC8SPjJQfHhRnfjQwRUeA8SBFatM")
db = deta.Base("resume")
slist={}
dlist={}
datascience_keywords=["Data","data","Analyst","analyst", "data scientist", "Data Scientist", "recommendation","machine learning","data cleaning"]
software_keywords=["web development","full stack","data structures","web","development","application","app","front end"]
if(choice=="Student"):
    st.header("RESUME")
    with st.form("form"):
        
        name=st.text_input("Your Name")
        gender=st.radio("Gender: ",('Male','Female'))
        course=st.text_input("Enter Qualification (Eg MSc DCS)")
        aboutme=st.text_area("Describe About Yourself")
        phone=st.number_input("Phone Number",step=0)
        mail=st.text_input("Enter Mail Address")
        option=st.selectbox("Roles",['Data Science','Software'])
        skills=st.multiselect("Skills",['Html','Css','Javascript','PHP','Python','C','C++','Java','JSwing','JSP','Servlet','Machine Learning','Deep Learning','NLP','Statistics'])
        cgpa=st.number_input("CGPA",max_value=10.00)
        interest=st.text_input("Areas of Interest")
        
        submitted = st.form_submit_button("Submit")
        if(submitted):
            stre=" "
            fullskill=stre.join(skills)
            if(name!="" and aboutme!="" and len(str(phone))==10 and len(skills)!=0 and cgpa!=0.00 and interest!=""):                                
                db.put({"Name": name, "gender": gender, "course": course, "aboutme": aboutme, "phone": phone, "mail": mail, "option": option, "skills": fullskill, "cgpa": cgpa, "interest": interest})
                st.success("Submitted Successfully")
                speak("Your resume is successfully submitted, you will get an email confirmation once you get shortlisted")
            else:
                stre=" "
                st.write(stre.join(skills))
                st.error("Please fill out the form :(")
                speak("Please fill out the form")

else:
    
    rolehr=st.selectbox("Roles",['Data Science','Software'])
    db_content = db.fetch().items
    i=0
    dcount=0
    scount=0
    while(i<len(db_content)):
        if(db_content[i]["option"]=="Data Science"):
            dcount=dcount+1
        else: 
            scount=scount+1
        i=i+1
    
    score=[]
    key=[]
    result=[]
    cgpa=[]
    i=0
    if(rolehr=="Data Science"):
        if(dcount>0):
            while(i<len(db_content)):
                
                    if(db_content[i]["option"]==rolehr):
                        
                        sen=db_content[i]["aboutme"]
                        cgpa.append(db_content[i]["cgpa"])
                        result.append(any(ele in sen for ele in datascience_keywords))
                        
                        res=TextBlob(sen)
                        key.append(db_content[i]["key"])
                        score.append(res.sentiment.polarity)
                        
                    i=i+1
            avg=sum(score)/len(score)
            
            i=0
            j=0
            while(i<len(score)):
            
                if(result[i]==True and cgpa[i]>8 and score[i]>=avg):
                    finalscore.append(score[i])
                    dlist[j]=(db.get(key[i]))
                    j=j+1
                i=i+1
            commonfunction(dlist)
        else:
            st.error(":( No DataScience Applicant")
        
        
    else:
        if(scount>0):
           while(i<len(db_content)):
            
                if(db_content[i]["option"]==rolehr):
                    
                    sen=db_content[i]["aboutme"]
                    cgpa.append(db_content[i]["cgpa"])
                    result.append(any(ele in sen for ele in software_keywords))
                    
                    res=TextBlob(sen)
                    key.append(db_content[i]["key"])
                    score.append(res.sentiment.polarity)
                    
                i=i+1
           avg=sum(score)/len(score)
           
           i=0
           j=0
           st.write(score)
           while(i<len(score)):
                
                if(result[i]==True and cgpa[i]>8 and score[i]>=avg):
                    finalscore.append(score[i])
                    slist[j]=(db.get(key[i]))
                    j=j+1
                i=i+1
           
           commonfunction(slist)
        else:
            st.error(":( No Software Applicant")


    
                        
        
        
    
