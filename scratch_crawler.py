import json
import requests

def create_scratch_json():

    url = "https://api.taiwanlottery.com/TLCAPIWeB/Instant/List"

    params = {
        "ScratchName": "",
        "Money": "",
        "PageNum": 1,
        "PageSize": 100
    }

    data = requests.get(
        url,
        params=params
    ).json()

    games = []

    for item in data["content"]["scratchListInfos"]:

        detail = requests.get(
            "https://api.taiwanlottery.com/TLCAPIWeB/Instant/Detail",
            params={
                "ScratchId": item["scratchId"]
            }
        ).json()["content"]

        games.append({
            "id": item["scratchId"],
            "name": item["scratchName"],
            "price": item["money"],
            "firstPrize": item["firstPrize"],
            "odds": item["oddsOfWinning"],
            "image": item["picPath"],
            "listingDate": detail["listingDate"][:10],
            "downDate": detail["downDate"][:10]
        })

    with open(
        "scratch.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            games,
            f,
            ensure_ascii=False,
            indent=4
        )

    print("scratch.json完成")