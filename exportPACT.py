import requests
import openpyxl

#задача выгрузить всю базу пользователей с пакта, так как другого способа выгрузки нет - делаем через API. Нужно
# перебирать по 50-100 пользователей, выгрузка выдаётся с переменной на следующую страницу. Подробнее инфа
# https://pact-im.github.io/api-doc/?shell#get-all-conversations
token_pact = 'token_here'  #token pact
next_page = ' '
count = 2
PACT_ID = [] #ID компаний через запятую в пакте по которым нужно сделать выгрузку

def getPACT(pact,next_page,count):# запрашиваем данные с пакта

    if next_page != ' ':
        response = requests.get(  # делаем запрос на вторую и последующие страницы через токен next_page
            'https://api.pact.im/p1/companies/{pact}/conversations?per=100&from={next_page}'.format( #sort_direction=desc - запрос с конца списка
                pact=pact, PACT_ID=PACT_ID, next_page=next_page),
                        headers={
                            'X-Private-Api-Token': token_pact})
        length = len(response.json().get('data').get('conversations'))
        print(length)
        conv = response.json().get('data').get('conversations')
        i = 0
        for i in range(length):
            sheet[count][0].value = conv[i].get('name')
            sheet[count][1].value = conv[i].get('sender_external_id')
            sheet[count][2].value = pact
            print(conv[i].get('sender_external_id') + ' ' + conv[i].get('name'), end='\n')
            i += 1
            count += 1
        else:
            print(str(length) + ' end')
        return response.json().get('data').get('next_page'),count
    else:
        response = requests.get( #делаем первый запрос и получаем переменную next_page для след. запросов
            'https://api.pact.im/p1/companies/{pact}/conversations?per=100'.format( #sort_direction=desc - запрос с конца списка
                pact=pact, PACT_ID=PACT_ID),
            headers={
                'X-Private-Api-Token': token_pact})
        print(response)
        length = len(response.json().get('data').get('conversations')) # берём количество диалогов
        print(length)
        conv = response.json().get('data').get('conversations')
        i = 0
        for i in range(length): #записываем в файл имя и номер телефона пользователя из диалога
            sheet[count][0].value = conv[i].get('name')
            sheet[count][1].value = conv[i].get('sender_external_id')
            sheet[count][2].value = pact
            print(conv[i].get('sender_external_id') + ' ' + conv[i].get('name'), end='\n')
            count += 1
            i += 1
        else:
            print(str(length) + ' end')
        return response.json().get('data').get('next_page'),count
i = 0
for pact in PACT_ID: #проходимся по каждой компании
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'Name'
    sheet['B1'] = 'phone'
    sheet['C1'] = 'PACT_ID'
    while next_page != None: # Если страница последняя - то она вернёт next_page = None
        s = getPACT(pact,next_page,count)
        next_page = s[0]
        count = s[1]
        print(str(i) + ' page' + ' next_page=' + str(next_page) + ' count' + str(count))
        book.save("test" + str(pact) + "new" + ".xlsx")
        i += 1
    next_page = ' '
    count = 2
    book.close()
