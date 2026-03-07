from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO

app = Flask(__name__)

ROSTER_FILE = "roster.xlsx"
ROSTER_INFO = "roster_info.txt"

LAST_TABLE = None
LAST_DATES = None


def load_roster():

    if os.path.exists(ROSTER_FILE):

        roster = pd.read_excel(ROSTER_FILE)

        roster["Agent ID"] = roster["Agent ID"].astype(str)

        roster["Shift"] = pd.to_datetime(
            roster["Shift"], errors="coerce"
        ).dt.time

        return roster

    return None


def process_login(file):

    df = pd.read_excel(file)

    df.columns = ["UserName", "Agent Name", "DateTime", "Event"]

    df = df[df["Event"] == "LOGIN"]

    df["UserName"] = df["UserName"].astype(str)

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

                shift_text = shift_time.strftime("%H:%M:%S")

                table[agent]["shift"] = shift_text

                shift_dt = pd.to_datetime(str(date) + " " + shift_text)

                grace = shift_dt + timedelta(minutes=5)

                if login > grace:

                    status = "late"
                    table[agent]["late"] += 1

        login_time = login.strftime("%H:%M:%S")

        if login_time == "00:00:00":

            login_time = ""

        table[agent]["days"][date] = {
            "time": login_time,
            "status": status
        }

    return table, dates


@app.route("/", methods=["GET", "POST"])
def index():

    global LAST_TABLE, LAST_DATES

    roster_time = ""

    if os.path.exists(ROSTER_INFO):

        roster_time = open(ROSTER_INFO).read()

    if request.method == "POST":

        if "loginfile" in request.files:

            file = request.files["loginfile"]

            if file.filename != "":

                LAST_TABLE, LAST_DATES = process_login(file)

    return render_template(
        "index.html",
        table=LAST_TABLE,
        dates=LAST_DATES,
        roster_time=roster_time
    )


@app.route("/upload_roster", methods=["POST"])
def upload_roster():

    roster = request.files["roster"]

    if roster.filename != "":

        roster.save(ROSTER_FILE)

        time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

        with open(ROSTER_INFO, "w") as f:

            f.write(time)

    return redirect("/")


@app.route("/delete_roster")
def delete_roster():

    if os.path.exists(ROSTER_FILE):

        os.remove(ROSTER_FILE)

    if os.path.exists(ROSTER_INFO):

        os.remove(ROSTER_INFO)

    return redirect("/")


@app.route("/reset")
def reset():

    global LAST_TABLE, LAST_DATES

    LAST_TABLE = None
    LAST_DATES = None

    return redirect("/")


@app.route("/export_excel")
def export_excel():

    rows = []

    for agent, data in LAST_TABLE.items():

        row = {
            "Agent": agent,
            "Name": data["name"],
            "Shift": data["shift"],
            "Late Count": data["late"]
        }

        for d in LAST_DATES:

            if d in data["days"]:

                row[str(d)] = data["days"][d]["time"]

            else:

                row[str(d)] = ""

        rows.append(row)

    df = pd.DataFrame(rows)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df.to_excel(writer, index=False)

    output.seek(0)

    return send_file(
        output,
        download_name="login_report.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
