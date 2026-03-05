from flask import Flask, render_template, request, redirect
import pandas as pd
import os
from datetime import timedelta

app = Flask(__name__)

ROSTER_FILE = "roster.xlsx"


# =============================
# LOAD ROSTER
# =============================
def load_roster():

    if os.path.exists(ROSTER_FILE):

        roster = pd.read_excel(ROSTER_FILE)

        roster["Shift"] = pd.to_datetime(
            roster["Shift"], errors="coerce"
        ).dt.time

        return roster

    return None


# =============================
# PROCESS LOGIN REPORT
# =============================
def process_login(file):

    df = pd.read_excel(file)

    df.columns = ["UserName", "Agent Name", "DateTime", "Event"]

    df = df[df["Event"] == "LOGIN"]

    df["DateTime"] = pd.to_datetime(df["DateTime"])

    df["Date"] = df["DateTime"].dt.date

    first_login = df.sort_values("DateTime").groupby(
        ["UserName", "Agent Name", "Date"]
    ).first().reset_index()

    roster = load_roster()

    table = {}

    dates = sorted(first_login["Date"].unique())

    for _, row in first_login.iterrows():

        agent = row["UserName"]
        name = row["Agent Name"]
        date = row["Date"]
        login = row["DateTime"]

        if agent not in table:

            table[agent] = {
                "name": name,
                "days": {},
                "late": 0,
                "shift": ""
            }

        shift_row = roster[roster["Agent ID"] == agent]

        status = ""

        if len(shift_row) > 0:

            shift_time = shift_row.iloc[0]["Shift"]

            if pd.notna(shift_time):

                table[agent]["shift"] = shift_time

                shift_dt = pd.to_datetime(str(date) + " " + str(shift_time))

                grace = shift_dt + timedelta(minutes=5)

                if login > grace:

                    status = "late"

                    table[agent]["late"] += 1

        table[agent]["days"][date] = {
            "time": login.strftime("%H:%M:%S"),
            "status": status
        }

    return table, dates


# =============================
# HOME PAGE
# =============================
@app.route("/", methods=["GET", "POST"])
def index():

    table = None
    dates = None

    if request.method == "POST":

        if "loginfile" in request.files:

            loginfile = request.files["loginfile"]

            if loginfile.filename != "":

                table, dates = process_login(loginfile)

    return render_template("index.html", table=table, dates=dates)


# =============================
# ROSTER UPLOAD
# =============================
@app.route("/upload_roster", methods=["POST"])
def upload_roster():

    roster = request.files["roster"]

    if roster.filename != "":

        roster.save(ROSTER_FILE)

    return redirect("/")


# =============================
# DELETE ROSTER
# =============================
@app.route("/delete_roster")
def delete_roster():

    if os.path.exists(ROSTER_FILE):

        os.remove(ROSTER_FILE)

    return redirect("/")


# =============================
# RUN SERVER
# =============================
if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
