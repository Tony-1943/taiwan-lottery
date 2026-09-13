import requests
from datetime import datetime
now = datetime.now()
# =========================
# 3星彩
# =========================
def get_3star():

    url = "https://api.taiwanlottery.com/TLCAPIWeB/Lottery/3DResult"

    params = {
        "month": f"{now.year}-{now.month-2:02d}",
        "endMonth": f"{now.year}-{now.month:02d}",
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
        "month": f"{now.year}-{now.month-2:02d}",
        "endMonth": f"{now.year}-{now.month:02d}",
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