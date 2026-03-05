from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import timedelta

app = Flask(__name__)

ROSTER_FILE = "roster.xlsx"

def load_roster():
    if os.path.exists(ROSTER_FILE):
        roster = pd.read_excel(ROSTER_FILE)
        roster["Shift"] = pd.to_datetime(roster["Shift"], format="%H:%M").dt.time
        return roster
    return None

def process_login(file):

    df = pd.read_excel(file)

    df.columns = ["UserName", "Agent Name", "DateTime", "Event"]

    df = df[df["Event"] == "LOGIN"]

    df["DateTime"] = pd.to_datetime(df["DateTime"])

    df["Date"] = df["DateTime"].dt.date
    df["Time"] = df["DateTime"].dt.time

    first_login = df.sort_values("DateTime").groupby(
        ["UserName", "Agent Name", "Date"]
    ).first().reset_index()

    roster = load_roster()

    result = []

    for _, row in first_login.iterrows():

        agent = row["UserName"]
        login_time = row["DateTime"]
        date = row["Date"]

        shift_row = roster[roster["Agent ID"] == agent]

        if len(shift_row) == 0:
            shift_time = None
            status = ""
        else:
            shift_time = shift_row.iloc[0]["Shift"]

            shift_datetime = pd.to_datetime(str(date) + " " + str(shift_time))

            grace = shift_datetime + timedelta(minutes=5)

            if login_time > grace:
                status = "late"
            else:
                status = "ok"

        result.append({
            "Agent": row["Agent Name"],
            "Date": date,
            "Day": pd.to_datetime(date).strftime("%A"),
            "Login": login_time.strftime("%H:%M:%S"),
            "Status": status
        })

    return pd.DataFrame(result)

@app.route("/", methods=["GET", "POST"])
def index():

    table = None

    if request.method == "POST":

        if "roster" in request.files:
            roster = request.files["roster"]
            if roster.filename != "":
                roster.save(ROSTER_FILE)

        if "loginfile" in request.files:

            loginfile = request.files["loginfile"]

            if loginfile.filename != "":
                df = process_login(loginfile)
                table = df.to_dict(orient="records")

    return render_template("index.html", table=table)

if __name__ == "__main__":
    app.run(debug=True)
