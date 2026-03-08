import requests
import pandas as pd


def fetch_pr_votes():

    session = requests.Session()

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    page_url = "https://result.election.gov.np/PRVoteChartResult2082.aspx"

    # Step 1 — get cookies
    page = session.get(page_url, headers=headers)

    cookies = session.cookies.get_dict()
    csrf_token = cookies.get("CsrfToken")

    # Step 2 — fetch JSON data
    json_url = "https://result.election.gov.np/Handlers/SecureJson.ashx?file=JSONFiles/Election2082/Common/PRHoRPartyTop5.txt"

    headers.update({
        "Referer": page_url,
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRF-Token": csrf_token
    })

    response = session.get(json_url, headers=headers)

    data = response.json()

    party_data = []
    total_votes = 0

    for row in data:

        party = row["PoliticalPartyName"]
        votes = int(row["TotalVoteReceived"])
        symbol = row["SymbolID"]

        logo = f"https://result.election.gov.np/Images/symbol-hor-pa/{symbol}.jpg"

        party_data.append({
            "Party": party,
            "Votes": votes,
            "Logo": logo
        })

        total_votes += votes

    df = pd.DataFrame(party_data)

    return df, total_votes