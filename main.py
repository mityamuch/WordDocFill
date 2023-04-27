import json
import locale
import os
from docx2pdf import convert
from docxtpl import DocxTemplate
import random
import pathlib


def makeExchangeCityTravelticket(usr):
    print(f'\r[+] Заполняю билет : {usr["name"]}', end='')

    data = {'tickN': usr["tickN"],
            'aviaN': usr["aviaN"],
            'orderN': usr["orderN"],
            'name': usr["name"],
            'bornDate': usr["bornDate"],
            'passportN': usr["passportN"],
            'timeFly': usr["timefly"],
            'planeN': usr["planeN"],
            'aviaComp': usr["aviaComp"],
            'planeType': usr["planeType"],
            'baggage': usr["baggage"],
            'outTime': usr["outTime"],
            'outDate': usr["outDate"],
            'regTime': usr["regtime"],
            'portNameS': usr["portNameS"],
            'city': usr["city"],
            'coutry': usr["country"],
            'portName': usr["portName"],
            'inTime': usr["inTime"],
            'inDate': usr['inDate'],
            'distance': usr['distance'],
            'orderDate': usr["orderDate"],
            'orderTime': usr['orderTime'],
            'sum': usr["sum"],
            'val': usr['val'],
            'planeN2': usr['planeN2'],
            'aviaComp2': usr['aviaComp2'],
            'planeType2': usr['planeType2'],
            'outDate2': usr['outDate2'],
            'outTime2': usr['outTime2'],
            'regTime2': usr['regTime2'],
            'portNameS2': usr['portNameS2'],
            'city2': usr['city2'],
            'coutry2': usr['coutry2'],
            'portName2': usr['portName2'],
            'inTime2': usr['inTime2'],
            'inDate2': usr['inDate2'],
            'distance2': usr['distance2'],
            'transferType': usr['transferType'],
            'portTransferName': usr['portTransferName'],
            'timeTransfer': usr['timeTransfer'],
            'aviaPhone': usr['aviaPhone'],
            }
    docticket = DocxTemplate("templateCityTravel2city.docx")
    docticket.render(data)
    docticket.save(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'))
    date = (usr["orderDate"] + usr["orderTime"]).replace(".", "").replace(':', '')
    filename = f'{usr["orderN"]}_{date}{random.randint(1000000, 9999999)}forward_aerovaucher.pdf'
    convert(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'),
            os.path.join(os.getcwd(), 'personal', usr["name"], filename))


def makeSimpleCityTravelTicket(usr):
    print(f'\r[+] Заполняю билет : {usr["name"]}', end='')

    data = {'tickN': usr["tickN"],
            'aviaN': usr["aviaN"],
            'orderN': usr["orderN"],
            'name': usr["name"],
            'bornDate': usr["bornDate"],
            'passportN': usr["passportN"],
            'timeFly': usr["timefly"],
            'planeN': usr["planeN"],
            'aviaComp': usr["aviaComp"],
            'planeType': usr["planeType"],
            'baggage': usr["baggage"],
            'outTime': usr["outTime"],
            'outDate': usr["outDate"],
            'regTime': usr["regtime"],
            'portNameS': usr["portNameS"],
            'city': usr["city"],
            'coutry': usr["country"],
            'portName': usr["portName"],
            'inTime': usr["inTime"],
            'inDate': usr['inDate'],
            'distance': usr['distance'],
            'orderDate': usr["orderDate"],
            'orderTime': usr['orderTime'],
            'sum': usr["sum"],
            'val': usr['val'],
            'aviaPhone': usr['aviaPhone'],
            }
    docticket = DocxTemplate("templateCitytravel.docx")
    docticket.render(data)
    docticket.save(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'))
    date = (usr["orderDate"] + usr["orderTime"]).replace(".", "").replace(':', '')
    filename = f'{usr["orderN"]}_{date}{random.randint(1000000, 9999999)}forward_aerovaucher.pdf'
    convert(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'),
            os.path.join(os.getcwd(), 'personal', usr["name"], filename))


def makeSimpleCityTravelTicketEn(usr):
    print(f'\r[+] Заполняю билет : {usr["name"]}', end='')

    data = {'tickN': usr["tickN"],
            'aviaN': usr["aviaN"],
            'orderN': usr["orderN"],
            'name': usr["name"],
            'bornDate': usr["bornDate"],
            'passportN': usr["passportN"],
            'timeFly': usr["timefly"],
            'planeN': usr["planeN"],
            'aviaComp': usr["aviaComp"],
            'planeType': usr["planeType"],
            'baggage': usr["baggage"],
            'outTime': usr["outTime"],
            'outDate': usr["outDate"],
            'regTime': usr["regtime"],
            'portNameS': usr["portNameS"],
            'city': usr["city"],
            'coutry': usr["country"],
            'portName': usr["portName"],
            'inTime': usr["inTime"],
            'inDate': usr['inDate'],
            'distance': usr['distance'],
            'orderDate': usr["orderDate"],
            'orderTime': usr['orderTime'],
            'sum': usr["sum"],
            'val': usr['val'],
            'aviaPhone': usr['aviaPhone'],
            }
    docticket = DocxTemplate("templateCitytravelEn.docx")
    docticket.render(data)
    docticket.save(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'))
    date = (usr["orderDate"] + usr["orderTime"]).replace(".", "").replace(':', '')
    filename = f'{usr["orderN"]}_{date}{random.randint(1000000, 9999999)}forward_aerovaucher.pdf'
    convert(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'),
            os.path.join(os.getcwd(), 'personal', usr["name"], filename))


def makeExchangeCityTravelticketEn(usr):
    print(f'\r[+] Заполняю билет : {usr["name"]}', end='')

    data = {'tickN': usr["tickN"],
            'aviaN': usr["aviaN"],
            'orderN': usr["orderN"],
            'name': usr["name"],
            'bornDate': usr["bornDate"],
            'passportN': usr["passportN"],
            'timeFly': usr["timefly"],
            'planeN': usr["planeN"],
            'aviaComp': usr["aviaComp"],
            'planeType': usr["planeType"],
            'baggage': usr["baggage"],
            'outTime': usr["outTime"],
            'outDate': usr["outDate"],
            'regTime': usr["regtime"],
            'portNameS': usr["portNameS"],
            'city': usr["city"],
            'coutry': usr["country"],
            'portName': usr["portName"],
            'inTime': usr["inTime"],
            'inDate': usr['inDate'],
            'distance': usr['distance'],
            'orderDate': usr["orderDate"],
            'orderTime': usr['orderTime'],
            'sum': usr["sum"],
            'val': usr['val'],
            'planeN2': usr['planeN2'],
            'aviaComp2': usr['aviaComp2'],
            'planeType2': usr['planeType2'],
            'outDate2': usr['outDate2'],
            'outTime2': usr['outTime2'],
            'regTime2': usr['regTime2'],
            'portNameS2': usr['portNameS2'],
            'city2': usr['city2'],
            'coutry2': usr['coutry2'],
            'portName2': usr['portName2'],
            'inTime2': usr['inTime2'],
            'inDate2': usr['inDate2'],
            'distance2': usr['distance2'],
            'transferType': usr['transferType'],
            'portTransferName': usr['portTransferName'],
            'timeTransfer': usr['timeTransfer'],
            'aviaPhone': usr['aviaPhone'],
            }
    docticket = DocxTemplate("templateCityTravel2city.docx")
    docticket.render(data)
    docticket.save(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'))
    date = (usr["orderDate"] + usr["orderTime"]).replace(".", "").replace(':', '')
    filename = f'{usr["orderN"]}_{date}{random.randint(1000000, 9999999)}forward_aerovaucher.pdf'
    convert(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["name"]}.docx'),
            os.path.join(os.getcwd(), 'personal', usr["name"], filename))


def makePolicy(usr):
    print(f'\r[+] Заполняю страховку : {usr["name"]}', end='')

    policydata = {
        'policN': usr["policN"],
        'policDate': usr["policDate"],
        'enName': usr["enName"],
        'rusName': usr["rusname"],
        'phone': usr["phone"],
        'bornDate': usr["bornDate"],
        'periodFrom': usr["periodFrom"],
        'periodTo': usr["periodTo"],
        'days': usr["days"],
        'randomEUR': usr["randomEUR"],
    }
    docpolice = DocxTemplate("SberPolice.docx")
    docpolice.render(policydata)
    docpolice.save(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["enName"]}_policy.docx'))

    filename = f'Polis_VZR_{usr["enName"]}.pdf'
    convert(os.path.join(os.getcwd(), 'personal', usr["name"], f'{usr["enName"]}_policy.docx'),
            os.path.join(os.getcwd(), 'personal', usr["name"], filename))


def makeBooking(usr):
    print(f'\r[+] Заполняю бронь жилья : {usr["name"]}', end='')

    policydata = {
        'bookN': usr["bookN"],
        'pin': usr["pin"],
        'inDate': usr["inDate"],
        'InMonth': usr["InMonth"],
        'inWeekday': usr["inWeekday"],
        'toDate': usr["toDate"],
        'toMonth': usr["toMonth"],
        'toWeekday': usr["toWeekday"],
        'night': usr["night"],
        'val': usr["bval"],
        'sum': usr["bsum"],
        'sumShek': usr["sumShek"],
        'USDval': usr["USDval"],
        'name': usr["enName"],
        'datePred': usr["datePred"],
        'inTimeArival': usr["inTimeArival"],
        'toTimeArival': usr["toTimeArival"],
        'mail': usr["mail"],
    }

    docbooking = DocxTemplate("Booking.docx")
    docbooking.render(policydata)
    docbooking.save(os.path.join(os.getcwd(), 'personal', usr["name"], f'Booking.com_ Подтверждение.docx'))

    filename = f'Booking.com_ Подтверждение.pdf'
    convert(os.path.join(os.getcwd(), 'personal', usr["name"], f'Booking.com_ Подтверждение.docx'),
            os.path.join(os.getcwd(), 'personal', usr["name"], filename))


def filling_docs():
    locale.setlocale(locale.LC_ALL, '')
    jsonpath = input("Введите путь к json файлу данных клиента: ")
    if not pathlib.Path(jsonpath).exists():
        jsonpath = pathlib.Path(jsonpath.strip('"'))

    if not pathlib.Path(jsonpath).exists():
        print("Ошибка: указанный файл не существует")
        return

    usr = json.load(open(jsonpath, 'r'))
    if not os.path.isdir(os.path.join(os.getcwd(), 'personal')):
        os.mkdir(os.path.join(os.getcwd(), 'personal'))
    if not os.path.isdir(os.path.join(os.getcwd(), 'personal', usr["name"])):
        os.mkdir(os.path.join(os.getcwd(), 'personal', usr["name"]))

    print(f'Выберите тип документов для формирования товарищу {usr["name"]}')

    print("1 - Билет на самолет")
    print("2 - Страховка")
    print("3 - Весь пакет документов")

    choice = input("Введите номер выбранного типа документов: ")
    if choice == "1":
        print("Выбран тип документов 1")

        print("Введите en для создание билета на английском языке")
        print("Введите ru для создание билета на английском языке")
        lang = input("Введите номер выбранного языка: ")
        if lang == "en":
            if usr["planeN2"] == "":
                print("Создаю билет без пересадки")
                makeSimpleCityTravelTicketEn(usr)
            else:
                print("Создаю билет с пересадкой")
                # makeExchangeCityTravelticket(usr)
        elif lang == "ru":
            if usr["planeN2"] == "":
                print("Создаю билет без пересадки")
                makeSimpleCityTravelTicket(usr)
            else:
                print("Создаю билет с пересадкой")
                makeExchangeCityTravelticket(usr)
        else:
            print("Ошибка: неверный выбор языка")


    elif choice == "2":
        print("Выбран тип документов 2")
        makePolicy(usr)

    elif choice == "3":
        print("Выбран тип документов 3")
        makeBooking(usr)

    elif choice == "3":
        print("Выбран тип документов 3")
        if usr["planeN2"] == "":
            print("Создаю билет без пересадки")
            makeSimpleCityTravelTicket(usr)
        else:
            print("Создаю билет с пересадкой")
            makeExchangeCityTravelticket(usr)
        makePolicy(usr)

    else:
        print("Ошибка: неверный выбор типа документов")


def main():
    filling_docs()
    print('\n[+] Все записи обработаны')
    os.system("pause")


if __name__ == "__main__":
    main()
