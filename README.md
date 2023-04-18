# Travel Documents Generator

Travel Documents Generator is a Python script designed to automate the process of generating travel documents, such as flight tickets, insurance policies, and hotel bookings, using provided data in JSON format. The script reads the input JSON file and fills the corresponding document templates with the data. The generated documents are then saved as PDF files in a specified directory.

Features
Generate flight tickets (with or without layovers)
Generate insurance policies
Generate hotel booking confirmations
Generate a complete package of all three documents

Requirements
Python 3.6 or higher
docx2pdf
docxtpl
python-docx

Usage
Place the JSON file containing the client's data in a suitable directory.
Run the script main.py:

python main.py

When prompted, enter the path to the JSON file:

```
Введите путь к json файлу данных клиента: path/to/your/json/file.json
```

Choose the type of documents to generate:
```
Выберите тип документов для формирования товарищу John Doe
1 - Билет на самолет
2 - Страховка
3 - Бронирование отеля
4 - Весь пакет документов
```

The script will generate the chosen documents and save them as PDF files in the /personal/ClientName directory.
JSON Data Format
The JSON file should contain the client's data in the following format:

```
json
{
  "name": "John Doe",
  "bornDate": "01.01.1980",
  "passportN": "123456789",
  "timeFly": "03:00",
  ...
}
```
Refer to the example JSON files provided in the repository for a complete list of fields.

Document Templates
This script uses the following document templates:

templateCityTravel2city.docx: Flight ticket with layovers
templateCitytravel.docx: Flight ticket without layovers
SberPolice.docx: Insurance policy
Booking.docx: Hotel booking confirmation
Customize these templates to match your desired output.
