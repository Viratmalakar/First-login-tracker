from flask import Flask, render_template, request
import pandas as pd
import json
import os
from datetime import timedelta

app = Flask(__name__)

AGENT_FILE = "agents.json"

TABLE = None
DATES = None


def load_agents():

    if os.path.exists(AGENT_FILE):

        with open(AGENT_FILE) as f:
            return json.load(f)

    return {}


def process(file):

    agents = load_agents()

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

    df["Day"] = df["DateTime"].dt.strftime("%a")

    first = df.sort_values("DateTime").groupby(
        ["UserName", "Date", "Day"]
    ).first().reset_index()

    table = {}

    dates = sorted(first["Date"].unique())

    days = {}

    for d in first[["Date", "Day"]].drop_duplicates().values:

        days[str(d[0])] = d[1]

    for _, r in first.iterrows():

        agent = str(r["UserName"])

        if agent not in agents:
            continue

        name = agents[agent]["name"]
        shift = agents[agent]["shift"]

        date = r["Date"]
        login = r["DateTime"]

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

        table[agent]["days"][str(date)] = {
            "time": login.strftime("%H:%M:%S"),
            "status": status
        }

    return table, dates, days


@app.route("/", methods=["GET", "POST"])
def index():

    global TABLE, DATES, DAYS

    TABLE = None
    DATES = None
    DAYS = None

    if request.method == "POST":

        file = request.files["file"]

        if file.filename != "":

            TABLE, DATES, DAYS = process(file)

    return render_template(
        "index.html",
        table=TABLE,
        dates=DATES,
        days=DAYS
    )


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
