from flask import Flask,render_template,request,redirect
import pandas as pd
import json
import os
from datetime import timedelta

app=Flask(__name__)

AGENT_FILE="agents.json"

TABLE=None
DATES=None


def load_agents():

    if os.path.exists(AGENT_FILE):

        with open(AGENT_FILE) as f:

            return json.load(f)

    return {}


def save_agents(data):

    with open(AGENT_FILE,"w") as f:

        json.dump(data,f,indent=4)



def process(file):

    agents=load_agents()

    df=pd.read_excel(file)

    df.columns=["UserName","Agent Name","DateTime","Event"]

    df=df[df["Event"]=="LOGIN"]

    df["UserName"]=df["UserName"].astype(str).str.replace(".0","",regex=False)

    df["DateTime"]=pd.to_datetime(df["DateTime"])

    df["Date"]=df["DateTime"].dt.date

    first=df.sort_values("DateTime").groupby(
        ["UserName","Date"]
    ).first().reset_index()

    table={}

    dates=sorted(first["Date"].unique())

    for _,r in first.iterrows():

        agent=str(r["UserName"])

        if agent not in agents:
            continue

        name=agents[agent]["name"]
        shift=agents[agent]["shift"]

        date=r["Date"]
        login=r["DateTime"]

        if agent not in table:

            table[agent]={
                "name":name,
                "shift":shift,
                "late":0,
                "days":{}
            }

        shift_dt=pd.to_datetime(str(date)+" "+shift)

        status=""

        if login>shift_dt+timedelta(minutes=5):

            status="late"
            table[agent]["late"]+=1

        table[agent]["days"][date]={
            "time":login.strftime("%H:%M:%S"),
            "status":status
        }

    return table,dates



@app.route("/",methods=["GET","POST"])

def index():

    global TABLE,DATES

    if request.method=="POST":

        file=request.files["file"]

        TABLE,DATES=process(file)

    return render_template("index.html",table=TABLE,dates=DATES)



@app.route("/settings")

def settings():

    agents=load_agents()

    return render_template("settings.html",agents=agents)



@app.route("/add_agent",methods=["POST"])

def add_agent():

    agents=load_agents()

    id=request.form["id"]

    name=request.form["name"]

    shift=request.form["shift"]

    agents[id]={"name":name,"shift":shift}

    save_agents(agents)

    return redirect("/settings")



@app.route("/bulk_upload",methods=["POST"])

def bulk_upload():

    agents=load_agents()

    file=request.files["file"]

    df=pd.read_excel(file)

    for _,r in df.iterrows():

        id=str(r["Agent ID"])

        name=r["Agent Name"]

        shift=str(r["Shift"])

        agents[id]={"name":name,"shift":shift}

    save_agents(agents)

    return redirect("/settings")



@app.route("/delete_agent")

def delete_agent():

    agents=load_agents()

    id=request.args.get("id")

    if id in agents:

        del agents[id]

    save_agents(agents)

    return redirect("/settings")



@app.route("/update_shift",methods=["POST"])

def update_shift():

    agents=load_agents()

    id=request.form["id"]

    shift=request.form["shift"]

    agents[id]["shift"]=shift

    save_agents(agents)

    return redirect("/settings")



if __name__=="__main__":

    port=int(os.environ.get("PORT",10000))

    app.run(host="0.0.0.0",port=port)
