import asyncio
import requests
from aiohttp import ClientSession
import csv


async def get(url, session):
    async with session.get(url) as response:
        return await response.text()


def get_total_page():
    url = "https://api.tba.org.tw/FindLawyer"
    data = {
        "metier": "",
        "name": "",
        "location": "",
        "isForeignLaw": "false",
        "enableExpiry": "false",
        "page": 1
    }
    response = requests.post(url, json=data)
    return response.json()['TotalPages']


def get_last_page_data(last_page_num):
    url = "https://api.tba.org.tw/FindLawyer"
    data = {
        "metier": "",
        "name": "",
        "location": "",
        "isForeignLaw": "false",
        "enableExpiry": "false",
        "page": last_page_num
    }
    response = requests.post(url, json=data)
    return response.json()


async def get_lawyer_data_by_id(session, lawyer_id):
    url = f"https://api.tba.org.tw/FindLawyer/{lawyer_id}"
    MAX_RETRIES = 3
    retries = 0
    print(f"正在處理id: {lawyer_id}")
    while retries < MAX_RETRIES:
        try:
            async with session.get(url) as response:
                response_text = await response.text()
                if not response_text:
                    return
                response_json = await response.json()  # Await the json() method
                return response_json
        except asyncio.TimeoutError:
            print(f"Request timed out. {lawyer_id} Retrying... ({retries}/{MAX_RETRIES})")
        print(f"Request failed after {MAX_RETRIES} retries.")
    return {}


async def scrape_lawyers(loop_amount):
    async with ClientSession() as session:
        tasks = []
        for i in range(1, loop_amount+1):
            task = asyncio.ensure_future(get_lawyer_data_by_id(session, i))
            tasks.append(task)
        lawyer_list = await asyncio.gather(*tasks)
        return lawyer_list


async def write_to_csv(filename, data):
    with open(filename, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)


def preprocess_lawyer_list(lawyer_list, remain_fields):

    lawyer_list = [elem for elem in lawyer_list if elem is not None]
    remain_fields_list = []
    for lawyer in lawyer_list:
        remain_fields_list.append({field: lawyer[field] for field in remain_fields})

    for lawyer in remain_fields_list:
        lawyer['SEX'] = '男' if lawyer['SEX'] else '女'

    return remain_fields_list


def verify_amount(last_page_data, lawyer_list):
    return (last_page_data['TotalPages'] - 1) * 10 + len(last_page_data) == len(lawyer_list)


async def main():
    total_page = get_total_page()
    print(total_page)
    last_page_data = get_last_page_data(total_page)
    print(last_page_data)
    loop_amount = last_page_data['Item'][len(last_page_data['Item']) - 1]['ID']
    remain_fields = ['NAME', 'SEX', 'BIRTHPLACE', 'EMAIL', 'CONAME', 'COADDRESS', 'COPHONE1', 'COFAX1']
    lawyer_list = await scrape_lawyers(loop_amount)
    lawyer_list = preprocess_lawyer_list(lawyer_list, remain_fields)
    print(verify_amount(last_page_data, lawyer_list))
    await write_to_csv("lawyers.csv", lawyer_list)



if __name__ == "__main__":
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())
