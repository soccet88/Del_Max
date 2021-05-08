import requests
from bs4 import BeautifulSoup
import openpyxl


def get_data(adres):
    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36"
    }

    wb = openpyxl.Workbook()
    sheet = wb.active

    sheet['A1'] = 'Название Товара'
    sheet['B1'] = 'Ссылка на товар'
    sheet['C1'] = 'Цена'
    sheet['D1'] = "Кол-во объявлений у продавца"
    sheet['E1'] = "Номер телефона"
    sheet['F1'] = "Добавили в закладки"
    sheet['G1'] = "Просмотров"

    row = 2
    a, b = int(input()), int(input())

    info_adress = adres.split("?")
    first_adres = info_adress[0]
    last_adres = info_adress[1]

    item_urls = []

    for i in range(a, b):

        url  = first_adres + str(i) + "?" + last_adres
        req = requests.get(url, headers)

        with open(f"projects{i}.html", "w", encoding="utf8") as file:
            file.write(req.text)
        with open(f"projects{i}.html", encoding="utf8"  ) as file:
            src = file.read()

        soup = BeautifulSoup(src, "lxml")
        articles = soup.find_all("li", class_="simpleAds")

        for article in articles:
            item_url = "https://www.skelbiu.lt" + article.find("a").get("href")
            item_urls.append(item_url)

    item_data_list = []

    for item_url in item_urls:
        req = requests.get(item_url, headers)
        project_name =  item_url.split("/")[-1]

        with open(f"data/{project_name}.html", "w", encoding="utf8") as file:
            file.write(req.text)

        with open(f"data/{project_name}.html", encoding="utf8") as file:
            src = file.read()

        soup = BeautifulSoup(src, "lxml")
        project_data = soup.find("div", class_="itemscope")

        try:
            item_name = project_data.find("h1", itemprop="name").text
        except Exception:
            item_name = "No Data"

        try:
            item_price = project_data.find("p", class_="price").text
        except Exception:
            item_price = "No Data"

        try:
            profile_orders = project_data.find("div", class_="profile-stats").text
        except Exception:
            profile_orders = "No Data"

        try:
            profile_number = project_data.find("div", class_="primary").text
        except Exception:
            profile_number = "No Data"

        try:
            remembered_item = project_data.find("span", {"id": "ad-bookmarks-count"}).text
        except Exception:
            remembered_item = "No Data"

        try:
            view_item = project_data.find("div", class_="block showed").text
        except Exception:
            view_item = "No Data"


        sheet[row][0].value = item_name.strip()
        sheet[row][1].value = item_url.strip()
        sheet[row][2].value = item_price.strip()
        try:
            sheet[row][3].value = int(profile_orders.split(" ")[-1].strip())
        except Exception as e:
            sheet[row][3].value = 0
        sheet[row][4].value = profile_number.strip()
        sheet[row][5].value = int(remembered_item.strip())
        sheet[row][6].value = view_item.strip()

        row += 1

        wb.save('result.xlsx')
        wb.close()


print("Vvedite ssilku: ")
adr = str(input())
get_data(adr)
