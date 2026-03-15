from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import timedelta

app = Flask(__name__)

ROSTER_FILE = "roster.xlsx"

TABLE = None
DATES = None


def load_roster():

    roster = pd.read_excel(ROSTER_FILE)

    roster["Agent ID"] = (
        roster["Agent ID"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    roster["Shift"] = roster["Shift"].astype(str)

    return roster


def process_login(file):

    roster = load_roster()

    df = pd.read_excel(file)

    df.columns = ["UserName", "Agent Name", "DateTime", "Event"]

    df = df[df["Event"] == "LOGIN"]

    df["UserName"] = (
        df["UserName"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    df["DateTime"] = pd.to_datetime(df["DateTime"])

    df["Date"] = df["DateTime"].dt.date

    first_login = df.sort_values("DateTime").groupby(
        ["UserName", "Date"]
    ).first().reset_index()

    table = {}

    dates = sorted(first_login["Date"].unique())

    for _, row in first_login.iterrows():

        agent = str(row["UserName"])
        date = row["Date"]
        login = row["DateTime"]

        roster_row = roster[roster["Agent ID"] == agent]

        if len(roster_row) == 0:
            continue

        name = roster_row.iloc[0]["Agent Name"]
        shift = roster_row.iloc[0]["Shift"]

        if agent not in table:

            table[agent] = {
                "name": name,
                "shift": shift,
                "late": 0,
                "days": {}
            }

        shift_dt = pd.to_datetime(str(date) + " " + shift)

        status = ""

        if login > shift_dt + timedelta(minutes=5):

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

    global TABLE, DATES

    if request.method == "POST":

        file = request.files["file"]

        if file.filename != "":

            TABLE, DATES = process_login(file)

    return render_template(
        "index.html",
        table=TABLE,
        dates=DATES
    )


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
