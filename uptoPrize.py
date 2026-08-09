def get_uptoPrize():
    """
    取得大樂透、威力彩頭獎累積金額
    """

    import requests

    url = "https://api.taiwanlottery.com/TLCAPIWeB/Lottery/uptoPrize"

    lotto_jackpot = ""
    power_jackpot = ""

    try:
        data = requests.get(url, timeout=10).json()

        for item in data["content"]["uptoPrizeList"]:

            if item["gameCode"] == 5118:
                lotto_jackpot = {
                    "amount": item["prize"],
                    "desc": item["uptoDesc"]
                }

            elif item["gameCode"] == 5134:
                power_jackpot = {
                    "amount": item["prize"],
                    "desc": item["uptoDesc"]
                }

    except Exception as e:
        print("取得頭獎資料失敗:", e)

    return {
        "lotto_jackpot": lotto_jackpot,
        "power_jackpot": power_jackpot
    }