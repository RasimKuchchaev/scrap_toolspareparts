import openpyxl
import requests
from bs4 import BeautifulSoup
import time
import csv

headers = {
    "accept": "*/*",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36"
}

count = 0
url_count =2

bookqwe = openpyxl.open("toolspareparts.xlsx",read_only=True)
sheetqwe = bookqwe.active

book = openpyxl.Workbook()
sheet = book.active
sheet["A1"] = "product_url"
sheet["B1"] = "name"
sheet["C1"] = "sku"
sheet["D1"] = "image_url"
sheet["E1"] = "image_name"
sheet["F1"] = "diagram_url"
sheet["G1"] = "diagram_name"
sheet["H1"] = "details"
sheet["I1"] = "diagram"
sheet["J1"] = "diagram_product"
sheet["K1"] = "product_name"
sheet["L1"] = "price"

for item in range(2, 22):
    url_excell = sheetqwe[f"A{item}"].value

    r = requests.get(url=url_excell,headers=headers)
    with open("index5.html", "w") as file:
        file.write(r.text)

    time.sleep(5)

    with open("index5.html", "r") as file:
        src1 = file.read()

    soup = BeautifulSoup(src1, "lxml")

    product_url = url_excell
    name = soup.find("h1", class_="page-title").text.strip()
    sku = soup.find("div", class_="product attribute sku").find(class_="value").text.strip()
    image_url = soup.find("div", class_="gallery-placeholder _block-content-loading").find("img").get("src").strip()
    image_name = image_url.split("/")[-1].strip()
    url_count +=1

    if image_name == "placeholder.jpg":
        with open("problem.csv", "a") as file:
            writer = csv.writer(file)
            writer.writerow(
                [url_count, url_excell]
            )
        continue

    if soup.find("span", class_="diagram__download"):
        diagram_url = soup.find("span", class_="diagram__download").find("a").get("href")
        diagram_name =diagram_url.split("/")[-1].strip()
    else:
        diagram_url = "None"
        diagram_name = "None"

    details = soup.find("div", class_="product attribute description").text.strip()
    data_table = soup.find("div", class_="table-wrapper grouped").find("tbody").find_all("tr")

    sheet[f"A{2+count}"] = product_url
    sheet[f"B{2+count}"] = name
    sheet[f"C{2+count}"] = sku
    sheet[f"D{2+count}"] = image_url
    sheet[f"E{2+count}"] = image_name
    sheet[f"F{2+count}"] = diagram_url
    sheet[f"G{2+count}"] = diagram_name
    sheet[f"H{2+count}"] = details

    for item in data_table:
        data = item.find_all("td")
        diagram = data[0].text
        diagram_product = data[1].text
        product_name = item.find(class_="col item").find(class_="product-item-name").text.strip()
        if item.find(class_="price-box price-final_price"):
            price = item.find(class_="col item").find(class_="price-box price-final_price").text.strip()
        else:
            price = "NO LONGER AVAILABLE"

        sheet[f"I{2+count}"] = diagram
        sheet[f"J{2+count}"] = diagram_product
        sheet[f"K{2+count}"] = product_name
        sheet[f"L{2+count}"] = price

        count += 1

    book.save("my_book.xlsx")
    book.close()

