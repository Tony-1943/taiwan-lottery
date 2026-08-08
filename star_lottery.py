import requests

# =========================
# 3星彩
# =========================
def get_3star():

    url = "https://api.taiwanlottery.com/TLCAPIWeB/Lottery/3DResult"

    params = {
        "month": "2026-06",
        "endMonth": "2026-08",
        "pageNum": 1,
        "pageSize": 200
    }

    response = requests.get(url, params=params)
    response.raise_for_status()

    data = response.json()

    records = data["content"]["lotto3DRes"][:10]

    result = []

    for item in records:
        result.append({
            "期別": item["period"],
            "開獎日期": item["lotteryDate"][:10],
            "獎號": item["drawNumberAppear"]
        })

    return result


# =========================
# 4星彩
# =========================
def get_4star():

    url = "https://api.taiwanlottery.com/TLCAPIWeB/Lottery/4DResult"

    params = {
        "month": "2026-06",
        "endMonth": "2026-08",
        "pageNum": 1,
        "pageSize": 200
    }

    response = requests.get(url, params=params)
    response.raise_for_status()

    data = response.json()

    records = data["content"]["lotto4DRes"][:10]

    result = []

    for item in records:
        result.append({
            "期別": item["period"],
            "開獎日期": item["lotteryDate"][:10],
            "獎號": item["drawNumberAppear"]
        })

    return result