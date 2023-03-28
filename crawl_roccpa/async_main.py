import asyncio
from aiohttp import ClientSession
from bs4 import BeautifulSoup
import csv
import time


async def get(url, session):
    async with session.get(url) as response:
        return await response.text()


async def scrape_company_list(session):
    company_list = []
    url = get_url(1)
    html = await get(url, session)
    soup = BeautifulSoup(html, "html.parser")
    page_info = soup.find('span', class_='page-txt').text
    total_page = int(page_info[page_info.find("/")+1:page_info.find("頁")])

    for page in range(1, total_page+1):
        url = get_url(page)
        html = await get(url, session)
        soup = BeautifulSoup(html, "html.parser")
        rows = soup.findAll("tr")
        for row in rows:
            tds = row.find_all("td")
            if not tds:
                continue
            serial_number = tds[0].text
            accountant = tds[1].text
            office_name = tds[2].text
            more_url = tds[3].a.get('href') if tds[3].a else ""
            company_list.append({"序號": serial_number, "會計師姓名": accountant, "事務所名稱": office_name, "more_url": more_url})
    return company_list


async def scrape_company_detail(session, company):
    if not company['more_url']:
        return {}
    url = f"https://www.roccpa.org.tw/member_search/{company['more_url']}"
    MAX_RETRIES = 3
    retries = 0
    while retries < MAX_RETRIES:
        try:
            html = await asyncio.wait_for(get(url, session), timeout=240)
            soup = BeautifulSoup(html, "html.parser")
            member_data_ol = soup.find('article', class_='memberdata').ol
            fields = ['事務所地址', '事務所電話', '事務所傳真', 'E-mail', '會籍編號']
            company_detail = {}
            for li in member_data_ol.findAll("li"):
                th = li.find('b').text
                if '會籍編號' in th:
                    td = li.find('div').text.strip()
                    company_detail['會籍編號'] = td
                if th in fields:
                    td = li.find('div').text.strip()
                    company_detail[th] = td
            return company_detail
        except asyncio.TimeoutError:
            print(f"Request timed out. {company['序號']} Retrying... ({retries}/{MAX_RETRIES})")
        print(f"Request failed after {MAX_RETRIES} retries.")
    return company_detail


async def scrape_companies():
    async with ClientSession() as session:
        company_list = await scrape_company_list(session)
        tasks = []
        for company in company_list:
            task = asyncio.ensure_future(scrape_company_detail(session, company))
            tasks.append(task)
        company_details = await asyncio.gather(*tasks)
        for i, company in enumerate(company_list):
            company.update(company_details[i])
        return company_list


def get_url(page):
    get_params = {
        "AntiToken": "ywXGmdmH5F%2bpXAENs%2fsVkk4k7htDZvu585GLDotEd68%3d",
        "location": "",
        "fields": "2",
        "keys": "聯合",
        "p": page
    }
    return f"https://www.roccpa.org.tw/member_search/list2{get_params_string(get_params)}"


def get_params_string(params):
    result = "?"
    for key, value in params.items():
        result += f"{key}={value}&"

    return result[:-1]


async def write_to_csv(filename, data):
    pre_process_member_data(data)
    with open(filename, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)


def pre_process_member_data(member_list):
    for member in member_list:
        member.pop('more_url')


async def main():
    start_time = time.time()

    companies = await scrape_companies()

    await write_to_csv("companies.csv", companies)

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Elapsed time: {elapsed_time:.2f} seconds")


if __name__ == "__main__":
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())