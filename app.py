from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os
from datetime import timedelta
from io import BytesIO

app = Flask(__name__)

ROSTER_FILE = "roster.xlsx"
LAST_TABLE = None
LAST_DATES = None


def load_roster():

    if os.path.exists(ROSTER_FILE):

        roster = pd.read_excel(ROSTER_FILE)

        roster["Shift"] = pd.to_datetime(
            roster["Shift"], errors="coerce"
        ).dt.time

        return roster

    return None


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

        if len(shift_row) > 0:

            shift_time = shift_row.iloc[0]["Shift"]

            if pd.notna(shift_time):

                table[agent]["shift"] = shift_time

                shift_dt = pd.to_datetime(str(date) + " " + str(shift_time))

                grace = shift_dt + timedelta(minutes=5)

                if login > grace:

                    status = "late"
                    table[agent]["late"] += 1

                else:

                    status = ""

            else:

                status = ""

        else:

            status = ""

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

    message = ""

    if request.method == "POST":

        if "loginfile" in request.files:

            file = request.files["loginfile"]

            if file.filename != "":

                table, dates = process_login(file)

                LAST_TABLE = table
                LAST_DATES = dates

    return render_template(
        "index.html",
        table=LAST_TABLE,
        dates=LAST_DATES,
        message=message
    )


@app.route("/upload_roster", methods=["POST"])
def upload_roster():

    roster = request.files["roster"]

    if roster.filename != "":

        roster.save(ROSTER_FILE)

    return redirect("/?msg=roster_uploaded")


@app.route("/export_excel")
def export_excel():

    global LAST_TABLE, LAST_DATES

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
