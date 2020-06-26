from flask import Flask,render_template,request,make_response,redirect,session
import pandas as pd
import numpy as np
import openpyxl

import os
import re

from codes import mail_the_letter

app=Flask(__name__)


@app.route("/")
def index():
    df=pd.read_excel('Test1.xlsx')
    listed_df=df.values
    return render_template('index.html',listed_df=listed_df)


@app.route("/mail_sent",methods=["POST"])
def mail_sent():
    try:
        try:
            message="Please Provide All the Details."
            pdffile=request.form['docfile']
            whom_to_send=request.form['whom_to_send']

            if not(os.path.isfile(pdffile)):
                message="Base File Not Found! Please Try Again"
                raise Exception('File Not Found')
        except:
            return render_template('index.html',message=message)

        mail_the_letter(pdffile,whom_to_send)
     
        message=""
        return render_template("sucess.html")
    except:
        return render_template("failure.html")



@app.route("/candidate_details")
def candidate_details():
    df=pd.read_excel('Test1.xlsx')
    listed_df=df.values
    return render_template('candidate_details.html',listed_df=listed_df)


@app.route("/add_candidates",methods=["GET","POST"])
def add_candidates():
    if request.method=="POST":
        try:
            message="Some Error Occured While Inserting! Please Provide the Complete Details."
            name=request.form['name']
            college_name=request.form['college_name']
            email=request.form['email']
            phone_no=request.form['phone_no']

            if len(phone_no)!= 10:
                message="Please Enter the 10 digit Phone Number"
                raise Exception('Invalid Phone Number')
            
            #Regular Expression To Check Email
            regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'

            if not(re.search(regex,email)):
                message="Please Enter the Valid Email Address"
                raise Exception('Invalid Email')

            workbook_obj = openpyxl.load_workbook('Test1.xlsx')
            sheet_obj = workbook_obj.active

            sheet_obj.append([name,college_name,phone_no,email])
            workbook_obj.save('Test1.xlsx')
        
            message="Candidate Added Sucessfully"
            return render_template('add_candidate.html',message)
        except:
            return render_template("add_candidate.html",message=message)
    return render_template('add_candidate.html')
    

@app.route("/remove_candidate",methods=["GET","POST"])
def remove_candidate():
    df=pd.read_excel('Test1.xlsx')
    listed_df=df.values
    if request.method=="POST":
        try:
            message="Some Error Occured! Please fill all Details."
            name=request.form['delete_name']
            remove_confirm=request.form['remove_confirm']
            if remove_confirm =="I am Sure":
                df=df[df["FullName"]!=name]
                df.to_excel("Test1.xlsx",index=False)
                message="Name Deleted Sucessfully"
                
                return redirect('candidate_details')

            else:
                message="Please Confirm to Delete."
                raise Exception("Not Confirmation")
            
        except:
            return render_template('remove_candidate.html',listed_df=listed_df,message=message)

    
    return render_template('remove_candidate.html',listed_df=listed_df)



if __name__=='__main__':
    app.run(debug=False)