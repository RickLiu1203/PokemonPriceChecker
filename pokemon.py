from bs4 import BeautifulSoup as bs
import requests
import xlwings as xw
import datetime

wb = xw.Book('pokemon.xlsm')
sheet = wb.sheets("COLLECTION")


def get_ungraded(name, number, cardset):
    url = f"https://www.pricecharting.com/game/pokemon-{cardset}/{name}-{number}"
    page = requests.get(url)
    soup = bs(page.content, "html.parser")
    try:
        ungraded = soup.find("span", class_="price js-price").text
        ungraded = ungraded.strip()[1:]
        return float(ungraded)
    except AttributeError:
        return "Not Found"

def rw_excel():
    sheet.range("I8").value = "Script Running"
    sheet.range("I8").font.color = (255,0,0)

    entries = sheet.range("I10").value
    entries = int(entries.split(" ")[2])

    for i in range(3, entries+3):
        if type(sheet.range(f"A{i}").value) == float:
            card_set = sheet[f"B{i}"].value
            if card_set != None:
                card_set = card_set.replace(" ", "-")
                card_name = sheet[f"C{i}"].value
                card_name = card_name.replace(" ", "-")
                if sheet[f"D{i}"].value == "Holo Rare":
                    card_name += "-Holo"
                card_num = sheet[f"E{i}"].value
                try:
                    card_num = int(card_num.split("/")[0])
                except ValueError:
                    card_num = card_num.split("/")[0]

                sheet[f"G{i}"].value = get_ungraded(card_name,card_num,card_set)
    
    current_time = datetime.datetime.now()
    current_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
    sheet["A1"].value = f"UPDATED: {current_time}"
    sheet.range("G:G").number_format = "$#,##0.00"
    
    sheet.range("I8").value = "Script Not Running"
    sheet.range("I8").font.color = (50,168,84)

    wb.save()

if __name__ == "__main__":
    xw.Book("pokemon.xlsm").set_mock_caller()
    rw_excel()
