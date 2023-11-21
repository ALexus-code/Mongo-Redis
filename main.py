import pymongo
from xsls2strings import xsls2jsonstrings

pars_doc = xsls2jsonstrings('стек юл организации.xlsx')
list_length = len(pars_doc)
enterprise_num = []
for i in range(list_length):
    num = 'ENTERPRICE' + str(i)
    enterprise_num.append(num)
d1 = zip(enterprise_num, pars_doc)
da = dict(d1)
#print(da)
# Create the client
client = pymongo.MongoClient('localhost', 27017)
list1 = []
# создание актуальной бд
# with client:
#     db = client.Enterprise
#     # Достать все
#     collection = db['company']
#     cursor = collection.find({})
#     for document in cursor:
#         # print(document)
#         my_file = open("Актуальная_база.json", "w+")
#         my_file.write('{' + str(document)[46:])
#         my_file.close()

    # Вставить один документ
    #print(db.company.insert_one(da))
    # поиск
#     dic = db.company.find({'ENTERPRICE92': '{"Название" : " СМК-50 ","ИНН" : " 3123160962 ","КПП" : " 312301001 ","Телефон" : " 22-18-93 ","www" : "  ","НеобходимостьФискализацииПлатежейОнлайнКасса" : " СМК-50 ","ГИС_ДАТА" : " 0 ","ГИС_ГУИД" : " 17.03.23 ","РФИОРуководителя" : " 325d6a44-c4cb-40ac-bc92-c771293ba684 ","ИДДляДокументооборота" : "  ","ТелефонКонт" : "  ","ТипПользователяВеб" : " 8-910-362-77-48||8-910-362-77-48 ","ДатаЛиквидации" : " None ","ДатаРегистрации" : "  ","Категория" : "  ","ФактАдресКладр" : "  ","АдресКладр" : " 0 ","АдресДетально" : "  ","ФактАдрес" : " 308015#31#обл. Белгородская###ул. Чичерина#50Б###643# ","ДатаОГРН" : "  ","ГруппировкаПП" : "  ","ЗарегистрировавшийОрган" : "  ","Факс" : " 0 ","ОГРН" : "  ","Часы работы паспортного" : "  ","Часы работы кассы" : " 1073123019969 ","Часы работы бухгалтерии" : "  ","ФИОРуководителя" : "  ","ФИОГлБухгалтера" : "  ","Сообщение" : " Гусев Александр Геннадьевич ","Режим работы" : "  ","Примечание" : " 308015|1|31|-1|23|-1|2140|дом|50Б|корпус||кв.|||Российская Федерация|| ","Отрасль" : "  ","ОтделККМ" : "  ","ОКДП" : " 0 ","ОКОНХ" : "  ","ОКВЭД" : "  ","КомСбор" : "  ","Доступ" : "  ","ТНаименование" : " 0 ","РНаименование" : "  ","ДНаименование" : "  ","Наименование" : "  ","ДолжностьРуководителя" : "  ","СтавкаНДС" : " ООО СМК-50 ","Вариант НДС" : " директор ","Бюджет" : " По-умолчанию ","email" : " Плательщик ","ЭнергетТел" : " 0 ","ЭнергетEmail" : " gag@bats.ru ","БухТел" : "  ","БухEmail" : "  ","РукТел" : "  ","РукEmail" : "  ","Адрес" : "  "}'})
#     for item in dic:
#         list1.append(item)
# print(pd.DataFrame(data=list1))


# db = client['SeriesDB']
# series_collection = db['series']
# def insert_document(collection, data):
#     """ Function to insert a document into a collection and
#     return the document's id.
#     """
#     return collection.insert_one(data).inserted_id

# pars_doc = xsls2jsonstrings('стек юл организации.xlsx')
# list_length = len(pars_doc)
# for i in range(list_length):
#     zhest = pars_doc[i].replace('"', '\"')
#     #print(zhest)
#     print(insert_document(series_collection, zhest))


# print(insert_document(series_collection, zhest))

# print(json_string)
# jsonstring = json_string.replace('"','\"')
# print(jsonstring)
# use Python's open() function to load a JSON file
# with open("full.json", 'r', encoding='utf-8') as json_data:
#     print("full.json TYPE:", type(json_data))
#list_length = len(pars_doc)
# for i in range(list_length):
#     print(pars_doc[i].replace(" \ ",'').replace("'",'"'))
#     bson_example = bson.encode(pars_doc[i].replace(" \ ",'').replace("'",'"')[41:])
    #print(bson_example)
    #db.Interprise.insert_one(pars_doc[i].replace(" \ ", '').replace("'", '"')).inserted_id
#print(list_length)
# make sure the string is a valid JSON object first
# try:
#     # use json.loads() to validate the string and create JSON dict
#     json_docs = json.loads(jsonstring)
#
#     # loads() method returns a Python dict
#     print ("json_docs TYPE:", type(json_docs))
#
#     # return a list of all of the JSON document keys
#     print ("MongoDB collections:", list(json_docs.keys()))
#
# except ValueError as error:
#     # quit the script if string is not a valid JSON
#     print ("json.loads() ValueError for BSON object:", error)
#     quit()

#bson_example = bson.encode({"enterprise": [{"Название": " Строитель-1 ", "ИНН": " 3123095583 ", "КПП": " 312301001 ", "Телефон": " 30-26-04 ", "www": "  ", "НеобходимостьФискализацииПлатежейОнлайнКасса": " Строитель-1 ", "ГИС_ДАТА": " 0 ", "ГИС_ГУИД": "  ", "РФИОРуководителя": "  ", "ИДДляДокументооборота": "  ", "ТелефонКонт": "  ", "ТипПользователяВеб": " 30-26-04 ", "ДатаЛиквидации": " None ", "ДатаРегистрации": "  ", "Категория": "  ", "ФактАдресКладр": "  ", "АдресКладр": " 0 ", "АдресДетально": "  ", "ФактАдрес": "  ", "ДатаОГРН": "  ", "ГруппировкаПП": "  ", "ЗарегистрировавшийОрган": "  ", "Факс": " 0 ", "ОГРН": "  ", "Часы работы паспортного": "  ", "Часы работы кассы": "  ", "Часы работы бухгалтерии": "  ", "ФИОРуководителя": "  ", "ФИОГлБухгалтера": "  ", "Сообщение": "  ", "Режим работы": "  ", "Примечание": " 308001|1|31|-1|23|-1|1767|дом||корпус||кв.|||Российская Федерация ", "Отрасль": "  ", "ОтделККМ": "  ", "ОКДП": " 0 ", "ОКОНХ": "  ", "ОКВЭД": "  ", "КомСбор": "  ", "Доступ": "  ", "ТНаименование": " 0 ", "РНаименование": "  ", "ДНаименование": "  ", "Наименование": "  ", "ДолжностьРуководителя": "  ", "СтавкаНДС": " ГСК Строитель-1 ", "Вариант НДС": "  ", "Бюджет": " По-умолчанию ", "email": " Плательщик ", "ЭнергетТел": " 0 ", "ЭнергетEmail": " 5282923 ", "БухТел": "  ", "БухEmail": "  ", "РукТел": "  ", "РукEmail": "  ", "Адрес": "  "}]})
#print(bson_example)
#json_ohran1 = json.dumps(jsonstring)
#json_ohran2 = json.loads(jsonstring)
#print(json_ohran1)
#print(json_ohran2)
# db.Interprise.insert_one(json_ohran1).inserted_id
# print(db.Interprise.insert_one(json_ohran1).inserted_id)