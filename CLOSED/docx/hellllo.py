from openpyxl import load_workbook
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from translator import translate_text
from num2words import num2words
from docxtpl import DocxTemplate
from random import randint
from openpyxl import load_workbook
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from petrovich.main import Petrovich
from petrovich.enums import Case, Gender
from num2words import num2words
from docxtpl import DocxTemplate
from random import randint

def get_sum_str(summ, currency):
    data = requests.get(f"https://summa-propisyu.ru/?summ={summ}&vat=20&val={CURRENCY_CODE[currency]}&sep=1")
    page = BeautifulSoup(data.text, "lxml")
    sum_str = page.find('textarea', {'id':'result1'}).text
    return sum_str


def get_datv_name(name):
    try:
        p = Petrovich()
        name = name.split(' ')
        cased_lname = p.lastname(name[0], Case.DATIVE, gender=Gender.MALE)
        cased_fname = p.firstname(name[1], Case.DATIVE, gender=Gender.MALE)
        return ' '.join([cased_lname, cased_fname])
    except: return ' '.join(name)


def get_short_name(name):
    try:
        name = name.split(' ')
        short_name = name[0] + ' ' + name[1][0].upper() + '.' + name[2][0].upper() + '.'
        return short_name
    except: return ' '.join(name)




CUR = {'RUB':['Российский рубль', '643', 'Рублей'],
 'USD':['Доллар США', '840', 'Долларов США'],
 'KGS':['Сом', '417', 'Сом'],
 'CYN':['Юань', '156', 'Юаней'], 
 'EUR':['Евро', '978', 'Евро'],}


UNITS = {'м метр':'006',
 'кг килограмм':'166',
 'л литр':'112',
 'шт штука':'796',
 'PC unit':'796'}

CURRENCY_CODE = {'usd':'2',
                 'rub':'0',
                 'euro':'3'}


class Documents:

    def __init__(self, currency, document, shipper, seller, powerOfAttorney, products, consignee, payer, other_data, shablons_dir='./shablons/', distance_dir='./done/'):
        self.shipper_addres = translate_text(f'{shipper["street"]} {shipper["city"]} {shipper["country"]}', target_language='ru')
        self.seller_addres = translate_text(f'{seller["street"]} {seller["city"]} {seller["country"]}', target_language='ru')
        self.consignee_addres = translate_text(f'{consignee["street"]} {consignee["city"]} {consignee["country"]}', target_language='ru')
        self.payer_addres = translate_text(f'{payer["street"]} {payer["city"]} {payer["country"]}', target_language='ru')
        self.shablons_dir = shablons_dir
        self.distance_dir = distance_dir
        self.currency = currency.lower()
        self.document = document
        self.shipper = shipper
        self.powerOfAttorney = powerOfAttorney
        self.products = products
        self.consignee = consignee
        self.payer = payer
        self.other_data = other_data
        self.documents_number = document['number']
        self.documents_date = document['date']
        self.powerOfAttorney_number = translate_text(f"{''.join([i[0] for i in self.powerOfAttorney['driver'].split(' ')])}{self.documents_date.replace('.', '')}", target_language='ru')
        self.seller = seller
        self.shipper_info = translate_text(f'{self.shipper["name"]}, {self.shipper_addres}, {self.shipper["tin"]}', target_language='ru')
        self.shipper_director_short_name = translate_text(get_short_name(self.shipper['director']), target_language='ru')
        self.consignee_info = translate_text(f'{self.consignee["name"]}, {self.consignee_addres}, {self.consignee["tin"]}', target_language='ru')
        self.payer_info = translate_text(f'{self.payer["name"]}, {self.payer_addres}, {self.payer["tin"]}', target_language='ru')
        self.total_sum = '{0:,.2f}'.format(float(sum([float(i[-1]) for i in self.products])), 2).replace(',', ' ').replace('.', ',')
        self.total_products = translate_text(f'{len(self.products)} ({num2words(len(self.products), lang="ru")})', target_language='ru')
        self.total_products_count = '{0:,}'.format(sum([i[0] for i in self.products])).replace(',', ' ')
        self.driver_short_name = translate_text(get_short_name(powerOfAttorney["driver"]), target_language='ru')
        self.consignee_director_short_name = translate_text(get_short_name(self.consignee['director']), target_language='ru')
        self.payer_director_short_name = translate_text(get_short_name(self.payer['director']), target_language='ru')
        self.seller_director_short_name = translate_text(get_short_name(self.seller['director']), target_language='ru')
        self.seller_info = translate_text(f'{self.seller["name"]}, {self.seller_addres}, {self.seller["tin"]}', target_language='ru')

    def generate_ttn(self):


        wb = load_workbook(filename = f'{self.shablons_dir}ТТН.xlsx')
        ws = wb.get_sheet_by_name('TDSheet')
        
        
        ws['CU4'] = self.documents_number # Номер ТТН
        ws['CT217'] = self.documents_number # Номер ТТН вторая страница
        ws['CU5'] = self.documents_date # Дата ТТН


        ws['P6'] = self.shipper_info # Грузоотрпавитель полное наименование организации, адрес, номер телефона, ИНН
        ws['F218'] = self.shipper_info
        ws['CU6'] = self.shipper["okpo"] # ОКПО Грузоотрпавитель
        
        ws['AM208'] = self.shipper_director_short_name
        ws['AM212'] = self.shipper_director_short_name
        ws['Y245'] = self.shipper_director_short_name
        ws['AM210'] = self.shipper_director_short_name

        ws['P8'] =  self.consignee_info # Грузополучатель полное наименование организации, адрес, номер телефона, ИНН 
        ws['CU7'] = self.consignee['okpo'] # ОКПО Грузополучатель
        
        ws['P10'] =  self.payer_info # Плательщик полное наименование организации, адрес, номер телефона, ИНН
        ws['CU9'] = self.payer['okpo'] # ОКПО Плательщик
        ws['J222'] = ws['P10'].value
        
        
        ws['AV228'] = self.consignee_addres
        
        ws['CH192'] = self.total_sum # Сумма итого
        ws['CH193'] = self.total_sum # Сумма итого
        ws['CV201'] = self.total_sum # Сумма итого
        ws['B206'] = get_sum_str(self.total_sum, self.currency) # Сумма итого прописью
        

        ws['H197'] = self.total_products #'количество наименований'
        ws['H195'] = self.total_products #'количество наименований'
        ws['J204'] = self.total_products #'количество наименований'

  
        ws['H200'] = f'{int(self.total_products_count.replace(" ", ""))//100} ({num2words(sum([int(i[0]) for i in self.products])//100, lang="ru")})'
        ws['BW192'] = int(self.total_products_count.replace(" ", ''))//100
        ws['BW193'] = int(self.total_products_count.replace(" ", ''))//100
        ws['AD192'] = self.total_products_count
        ws['AD193'] = self.total_products_count


        ws['BD203'] = f'{get_datv_name(self.powerOfAttorney["driver"])} от {self.shipper["short_name"]}' 
        ws['CL208'] = self.driver_short_name
        ws['F224'] = self.driver_short_name

        ws['CC202'] = self.documents_date
        ws['BK202'] = self.powerOfAttorney_number
        ws['CL210'] = self.consignee_director_short_name

        for index, product in enumerate(self.products):
            index += 15
            ws[f'AD{index}'] = '{0:,}'.format(int(product[0])).replace(',', ' ')
            ws[f'AK{index}'] = '{0:,.2f}'.format(float(product[1]), 2).replace(',', ' ').replace('.', ',')
            ws[f'AR{index}'] = product[2]
            ws[f'BG{index}'] = product[3]
            ws[f'BM{index}'] = product[4]
            ws[f'BW{index}'] = '{0:,}'.format(int(product[0])//100).replace(',', ' ')
            ws[f'CH{index}'] = '{0:,.2f}'.format(float(product[5]), 2).replace(',', ' ').replace('.', ',')
            
        ws['F220'] = self.powerOfAttorney['car']
        ws['BH220'] = self.powerOfAttorney['car_number']
        ws['AV224'] = self.powerOfAttorney['prava']
        ws['F228'] = self.shipper_addres

        wb.save(self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+'ТТН.xlsx')


    def generate_upd(self):

        wb = load_workbook(filename = f'{self.shablons_dir}УПД.xlsx')
        ws = wb.get_sheet_by_name('TDSheet')    

        ws['P1'] = self.documents_number
        ws['Y1'] = self.documents_date

        ws['R4'] = self.seller['name']
        ws['R5'] = self.seller_addres
        ws['R6'] = self.seller['tin']

        ws['R7'] = self.shipper_info
        ws['R8'] = self.consignee_info


        ws['BE4'] = self.payer['name']
        ws['BE5'] = self.payer_addres
        ws['BE6'] = self.payer['tin']

        ws['AC194'] = self.payer_director_short_name

        ws['BN194'] = self.payer_director_short_name

        ws['Z205'] = self.shipper_director_short_name

        ws['O207'] = self.documents_date

        ws['BM205'] = self.consignee_director_short_name

        ws['Z213'] = self.seller_director_short_name

        ws['BM213'] = self.payer_director_short_name

        ws['C216'] = self.seller_info

        ws['AS216'] = self.payer_info

        ws['BE7'] = ', '.join(CUR[self.currency.upper()][:-1])

        ws['R10'] = f'ТТН №{self.documents_number} от {self.documents_date}'

        for index, product in enumerate(self.products):
            index += 15
            ws[f'J{index}'] = product[2]
            unit = [i for i in UNITS if product[3] in i][0]
            ws[f'Y{index}'] = unit.split(' ')[0]
            ws[f'W{index}'] = UNITS[unit]
            ws[f'AA{index}'] = '{0:,}'.format(int(product[0])).replace(',', ' ')
            ws[f'AD{index}'] = '{0:,.2f}'.format(float(product[1]), 2).replace(',', ' ').replace('.', ',')
            ws[f'AN{index}'] = '{0:,.2f}'.format(float(product[5]), 2).replace(',', ' ').replace('.', ',')
            ws[f'BF{index}'] = '{0:,.2f}'.format(float(product[5]), 2).replace(',', ' ').replace('.', ',')


        ws['AN192'] = self.total_sum
        ws['BF192'] = self.total_sum

        ws['O201'] = f'ТТН №{self.documents_number} от {self.documents_date}, {self.powerOfAttorney_number} от {self.documents_date}'

        wb.save(self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+'УПД.xlsx')


    def generate_cmr(self):
        cmr_data = {'line_2':translate_text(f"ТТН №{self.documents_number} от {self.documents_date}, УПД №{self.documents_number} от {self.documents_date}"),
            'line_1': translate_text(f'Контракт №{self.document["contract"]} от {self.document["contract_date"]}, Приложение №{self.document["appendix"]} от {self.document["appendix_date"]}'),
            'consignee_country':translate_text(self.consignee['country']),
            'consignee_street':translate_text(self.consignee['street']),
            'shipper_country':translate_text(self.shipper['country']),
            'shipper_name':translate_text(self.shipper['name']),
            'driver_name':translate_text(self.powerOfAttorney['driver']),
            'shipper_city':translate_text(self.shipper['city']),
            'documents_date':translate_text(self.documents_date),
            'driver_short_name':translate_text(self.driver_short_name),
            'car':translate_text(self.powerOfAttorney['car']),
            'car_number':translate_text(self.powerOfAttorney['car_number']),
            'consignee_name':translate_text(self.consignee['name']),
            'driver_passport':translate_text(self.powerOfAttorney['passport']),
            'currency_name':translate_text(CUR[self.currency.upper()][-1]),
            'driver_passport_date':translate_text(self.powerOfAttorney['passport_date']),
            'consignee_city':translate_text(self.consignee['city']),
            'shipper_street':translate_text(self.shipper['street']),
            'documents_number':translate_text(self.documents_number),}
        k = 0
        mal = 0
        cmr_number = randint(1000000, 9999999)
        mal_sum = 0
    
        cmr_data['product_1']=translate_text(f'{int(self.products[0][0])//100} {self.products[0][4]} {self.products[0][2]}')
        # for ind in range(len(self.products)):
        #     mal += 1
        #     mal_sum += float(self.products[ind][5])
        #     if ind%6==0:
        #         cmr_data['total']=('{0:,.2f}'.format(float(mal_sum), 2).replace(',', ' ').replace('.', ','))
        #         doc = DocxTemplate(f"{self.shablons_dir}CMR-1 .docx")
        #         cmr_data['cmr_number']=cmr_number+k
        #         inner = self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+f'CMR-{k}.docx'
        #         doc.render(cmr_data)
        #         doc.save(inner)
        #         k+=1
        #         mal = 0
        #         mal_sum = 0
        #         cmr_data[f'product_1'] = ''
        #         cmr_data[f'product_2'] = ''
        #         cmr_data[f'product_3'] = ''
        #         cmr_data[f'product_4'] = ''
        #         cmr_data[f'product_5'] = ''
        #         cmr_data[f'product_6'] = ''
        cmr_data['total']='{0:,.2f}'.format(float(mal_sum), 2).replace(',', ' ').replace('.', ',')
        doc = DocxTemplate(f"{self.shablons_dir}CMR-1 .docx")
        cmr_data['cmr_number']=cmr_number+k
        inner = self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+f'CMR-{k}.docx'
        doc.render(cmr_data)
        doc.save(inner)

    def generate_dover(self):
        
        data = {
            "company_name":self.payer['name'],
            "company_address":self.payer_addres,
            "company_tin":self.payer['tin'],
            "doc_number":self.powerOfAttorney_number,
            "doc_date":self.documents_date,
            "documents_date":self.documents_date,
            "driver_datv":get_datv_name(self.powerOfAttorney["driver"]),
            "passport":self.powerOfAttorney['passport'],
            "passport_date":self.powerOfAttorney['passport_date'],
            "company_short_name":self.payer['short_name'],
            "documents_number":self.documents_number,
            "consignee":self.consignee["name"],
            "consignee_address":self.consignee_addres,
            "director_short":self.payer_director_short_name,
            'country_consignee':self.consignee['country'],
            'country_shipper':self.shipper['country'],
            'city_Consignee':self.consignee['city'],
            'city_shipper':self.shipper['city'],
        }
        doc = DocxTemplate(f'{self.shablons_dir}Dover.docx')
        inner = self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+f'Доверенность.docx'
        doc.render(data)
        doc.save(inner)

    def gnerate_transport_request(self):
        products = []
        for product in self.products:
            pr = {
                'product_name_ru':product[2],
                'product_price':'{0:,.2f}'.format(float(product[5]), 2).replace(',', ' ').replace('.', ','),
                'product_price_str':get_sum_str(product[5], self.currency),
            }
            products.append(pr)
        order_date = {'car':self.powerOfAttorney['car'],
                      'payer_address':self.payer_addres,
                      'driver_full_name':self.powerOfAttorney['driver'],
                      'date_consignee':self.other_data['date_consignee'],
                      'payer_name':self.payer['name'],
                      'address_shipper':self.shipper_addres,
                      'products':products,
                      'shipper_company':self.shipper['name'],
                      'product_price_total':f'{self.total_sum} {CUR[self.currency.upper()][-1]} ({get_sum_str(self.total_sum, self.currency)})',
                      'country_consignee':self.consignee['country'],
                      'country_shipper':self.shipper['country'],
                      'city_Consignee':self.consignee['city'],
                      'city_shipper':self.shipper['city'],
                      'date_shipper':self.other_data['date_shipper'],
                      'payer_tin':self.payer['tin'],
                      'address_consignee':self.consignee_addres,
                      'car_number':self.powerOfAttorney['car_number'],
                      'documents_number':self.documents_number,
                      'documents_date':self.documents_date,
                      'consignee_company':self.consignee['name'],
                      'pll_number':self.other_data['pll_number']
                      }
        doc = DocxTemplate(f"{self.shablons_dir}Транспортная-заявка.docx")
        inner = self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+f'Transport requests.docx'
        doc.render(order_date)
        doc.save(inner)

    def generate_packing_list(self):
        products = []
        x = 1
        for product in self.products:
            pr = {
                'number':x,
                'name':product[2],
                'qtt':'{0:,}'.format(int(product[0])).replace(',', ' '),
                'unit':'{0:,.2f}'.format(float(product[1]), 2).replace(',', ' ').replace('.', ','),
                'total':'{0:,.2f}'.format(float(product[5]), 2).replace(',', ' ').replace('.', ',')
            }
            x+=1
            products.append(pr)
        data = {
            'total':f'{self.total_sum} {CUR[self.currency.upper()][-1]} ({get_sum_str(self.total_sum, self.currency)})', 
            'consignee_name':self.consignee['name'], 
            'contract_date':self.document['contract_date'], 
            'documents_date':self.documents_date, 
            'shipper_address':self.shipper_addres, 
            'contract_number':self.document['contract'], 
            'appendix_number':self.document['appendix'], 
            'shipper_name':self.shipper['name'], 
            'consignee_address':self.consignee_addres, 
            'appendix_date':self.document['appendix_date'], 
            'products':products, 
            'director_short_name':self.shipper_director_short_name, 
            'currency':CUR[self.currency.upper()][0],
            'documents_number':self.documents_number
            }
        doc = DocxTemplate(f"{self.shablons_dir}packing_list.docx")
        inner = self.distance_dir+str(datetime.now().timestamp()).split('.')[0]+f'Packing list.docx'
        doc.render(data)
        doc.save(inner)


    def generate(self):
        self.generate_ttn()
        print('ttn')
        # self.generate_upd() 
        self.generate_cmr()
        print('cmr')
        self.gnerate_transport_request()
        print('ransport_request')
        self.generate_packing_list()
        print('packing_list')
        self.generate_dover()
        print('dover')






pollllllllllll = {'powerOfAttorney1' : {
    'driver':'Оморов Бакыт Уланович', # выбор водителя рандом из базы
    'car':'Kamaz 5460', # машина привязання к водителю из базы
    'car_number':'O0498BB198', # номер машины привязанные к водителю из базы
    'prava':'  ', # права водителя привязанные к водителю из базы
    'passport':'ID2185347', # паспорт водителя из базы
    'passport_date':'МКК 211011 from 01.09.2019' # дата выдачи паспорта водителя из базы
},
'powerOfAttorney2' : {
    'driver':'Абдырахманов Эламан Тойгонбаевич', # выбор водителя рандом из базы
    'car':'DAF XF', # машина привязання к водителю из базы
    'car_number':'B5803AF', # номер машины привязанные к водителю из базы
    'prava':'  ', # права водителя привязанные к водителю из базы
    'passport':'AC4412412', # паспорт водителя из базы
    'passport_date':'SRS from 03.12.2020' # дата выдачи паспорта водителя из базы
},
'powerOfAttorney3' : {
    'driver':'Болинов Рамазан Юсупович', # выбор водителя рандом из базы
    'car':'KAMAZ 5460', # машина привязання к водителю из базы
    'car_number':'B102KC16', # номер машины привязанные к водителю из базы
    'prava':'  ', # права водителя привязанные к водителю из базы
    'passport':'ID0275632', # паспорт водителя из базы
    'passport_date':'МКК 218071 from 23.09.2017' # дата выдачи паспорта водителя из базы
},
'powerOfAttorney4' : {
    'driver':'Жакыпов Нурлан Азизович', # выбор водителя рандом из базы
    'car':'SHACMAN L3000', # машина привязання к водителю из базы
    'car_number':'CE9144', # номер машины привязанные к водителю из базы
    'prava':'  ', # права водителя привязанные к водителю из базы
    'passport':'AC4728463', # паспорт водителя из базы
    'passport_date':'SRS from 12.05.2019' # дата выдачи паспорта водителя из базы
},

'powerOfAttorney5' : {
    'driver':'Оморов Бакыт Уланович', # выбор водителя рандом из базы
    'car':'Kamaz 5460', # машина привязання к водителю из базы
    'car_number':'O0498BB198', # номер машины привязанные к водителю из базы
    'prava':'  ', # права водителя привязанные к водителю из базы
    'passport':'ID2185347', # паспорт водителя из базы
    'passport_date':'МКК 211011 from 01.09.2019' # дата выдачи паспорта водителя из базы
},}


ConsigneeNew= { #Грузопоучатель
    'name':'ООО НЬЮ-ЭКСИМ Пауэр',
    'short_name':'ООО НЬЮ-ЭКСИМ Пауэр',
    'okpo':'1095313',
    'director':'Кахаров Шохзод ЭркинОгли',
    'street':'ул. Мусаффо д-7.',
    'city':'Кувасой',
    'country':'Республика Узбекистан',
    'tin':'309267239',
}

ConsigneeKelich= { #Грузопоучатель
    'name':'КИЛИКАСЛАН ГИЙИМ ТЕКСТИЛЬ ГИДАКИЙМЕТЛИ МАДЕНЛЕР МАКИНА САНАЙИ ДИС ТИКАРЕТ ЛИМИТЕД СИРКЕТИ',
    'short_name':'КИЛИКАСЛАН ГИЙИМ ТЕКСТИЛЬ ГИДАКИЙМЕТЛИ МАДЕНЛЕР МАКИНА САНАЙИ ДИС ТИКАРЕТ ЛИМИТЕД СИРКЕТИ',
    'okpo':' ',
    'director':'Омер Килиджаслан',
    'street':'БЕЯЗИТ МАХ. ЯГЛИКЧИЛАР САД. ЦУКУРАН НЕТ. 63/5 КАПИ №. ФАТИХ',
    'city':'СТАМБУЛ',
    'country':'Турция',
    'tin':' ',
}

ConsigneePetro= { #Грузопоучатель
    'name': 'ООО "Петроимпэкс"',
    'short_name': 'ООО "Петроимпэкс"',
    'okpo': '654463',
    'director': 'ВАХИДОВ АЛИШЕР МИРОДИЛОВИЧ',
    'street': 'Нукусская улица 29, Мирабадский р-н.',
    'city': 'Ташкент',
    'country': 'Республика Узбекистан',
    'tin': '305888426',
}

ConsigneeOazr= { #Грузопоучатель
    'name':"ОсОО «СТРОЙКОНТ»",
    'short_name':"ОсОО «СТРОЙКОНТ»",
    'okpo':'30835446',
    'director':'Лебедев Андрей Юрьевич',
    'street':'ул. Токтогула 75/1',
    'city':'г. Бишкек',
    'country':'Кыргызская Республика',
    'tin':'41505202010178',
}

ConsigneeDemsa= { #Грузопоучатель
    'name':"ООО «СТРОЙКОНТ»",
    'short_name':"ООО «СТРОЙКОНТ»",
    'okpo':'39656018',
    'director':'Лебедев Андрей Юрьевич',
    'street':'Николоямская ул, д. 34 к. 2, помещ. 3/1',
    'city':'г. Москва',
    'country':'Россия',
    'tin':'9709049620',
}

Payer = { #Покупатель
    'name':'Общество с ограниченной ответственностью "Мастер Транс Компани"',
    'short_name':'ОсОО "Мастер Транс КейДжи"',
    'okpo':'31564989',
    'director':'Сагынова Нурзат Жанышовна',
    'street':'ул. Гоголь 9/90',
    'city':'Бишкек',
    'country':'Кыргызская Республика',
    'tin':'00709202210014',
}



import openai
import time
from random import randint
from faker import Faker
fake = Faker()

from datetime import datetime, date, timedelta
start_date = date(year=2023, month=5, day=20)
end_date = date(year=2023, month=5, day=25)


[12,18], [18, 60], [60, 91], [91, 121], [121, 139], [139, 166], [166, 182], [182, 187], [187, 191], [191, 197], [197, 206], [206, 218], [218, 227], [227, 230], [230, 239], [239, 247], [247, 248], [248, 251], [251, 257], [257, 275], [275, 282], [282, 295], [295, 307], [307, 324], [324, 345], [345, 348], [348, 350], [350, 353], [353, 354], [354, 356], [356, 360], [360, 362]
lm = [[12,18], [18, 60], [60, 91], [91, 121], [121, 139], [139, 166], [166, 182], [182, 187]]
lm = [[182, 187],] #Demsa
lm = [[230, 239],] #Kelich
lm = [[251, 257],] #New
lm = [[345, 348],] #Petro
lm = [[360, 362],] #Oazr
lm = [[12, 21],] #Oazr

# [182, 187], [230, 239], [251, 257], [345, 348],
# lm = [ [360, 362],]
wb = load_workbook(filename = 'VIPData.xlsx')
ws = wb.get_sheet_by_name('Лист1')
pppl = 1
isum = 0
x = 1
# ConsigneeDemsa, ConsigneeKelich, ConsigneeNew, ConsigneePetro,
a = [ ConsigneeOazr]

for i in lm:
    products = []
    for k in range(*i):
        # print([i for i in ws[f'O{k}'].value.split(',') if 'city' in i.lower()])
        looool = fake.date_between(start_date=start_date, end_date=end_date).strftime("%d.%m.%Y")
        print(ws[f"B{k}"].value, ws[f"G{k}"].value)
        Document = {
        'number':k,# вбивается вручную
        'date':looool, # вбивается вручную
        'contract':ws[f"B{k}"].value +'/'+ ws[f"G{k}"].value, # вбивается вручную
        'contract_date':'14.04.2023/'+str(ws[f"D{k}"].value), # вбивается вручную
        'appendix':ws[f"C{k}"].value, # вбивается вручную
        'appendix_date':str(ws[f"D{k}"].value), # вбивается вручную
        }
        pr = {
            'number':x,
            'name':ws[f'H{k}'].value.replace('\n', ''),
            'name_en':ws[f'H{k}'].value,
            'qtt':f"{ws[f'I{k}'].value}".replace(',', '.').replace(' ', ''),
            'unit':'{0:,.2f}'.format(float(f"{ws[f'J{k}'].value}".replace(',', '.').replace(' ', '')), 2),
            'total':'{0:,.2f}'.format(float(f"{ws[f'E{k}'].value}".replace(',', '.').replace(' ', '')), 2),
            'xunit':'{0:,.2f}'.format(float(f"{ws[f'E{k}'].value}".replace(',', '.').replace(' ', ''))*0.005, 2),
            'numberX':x+1,
        }
 
        isum+=float(f"{ws[f'E{k}'].value}".replace(',', '.').replace(' ', ''))
        x+=2
        isum+=float(float(f"{ws[f'E{k}'].value}".replace(',', '.').replace(' ', ''))*0.005,)
        products.append(pr)
        print(ws[f'O{k}'].value)
        # print the chat completion
        print(f"найди страну и город {ws[f'O{k}'].value}, ответ в виде Страна, Город, адрес на английском")
        names = ws[f"N{k}"].value
        Seller = { #Продавец
        'name':ws[f"N{k}"].value,
        'short_name':ws[f"N{k}"].value,
        'okpo':' ', 
        'director':ws[f"K{k}"].value, 
        'street':ws[f"O{k}"].value,
        'city':ws[f"Q{k}"].value,
        'country':ws[f"P{k}"].value,
        'tin': ' ', 
        }
        Shipper = { #Продавец
        'name':ws[f"N{k}"].value,
        'short_name':ws[f"N{k}"].value,
        'okpo':' ', 
        'director':ws[f"K{k}"].value, 
        'street':ws[f"O{k}"].value,
        'city':ws[f"Q{k}"].value,
        'country':ws[f"P{k}"].value,
        'tin': ' ', 
        }
        klll = datetime.strptime(looool, '%d.%m.%Y')+timedelta(days=18)
        other_data = {
            'date_consignee':klll.strftime("%d.%m.%Y"),
            'date_shipper':looool,
            'pll_number':ws[f"S{k}"].value
        }
        productsssss = [[int(ws[f'I{k}'].value), float(ws[f'J{k}'].value), ws[f'H{k}'].value, 'PC', 'corrugated box' ,float(f"{ws[f'E{k}'].value}".replace(',', '.').replace(' ', ''))],]
        docs = Documents(currency='usd', document=Document, shipper=Shipper, seller=Payer, powerOfAttorney=pollllllllllll[f"powerOfAttorney{randint(1,5)}"], products=productsssss, consignee=ConsigneeDemsa, payer=ConsigneeOazr, other_data=other_data)
        docs.generate()
        # time.sleep(10)
    
    data = {
        "appendix_number":ws[f"C{k}"].value,
        "addendix_date":datetime.strptime(ws[f"D{k}"].value, '%Y-%m-%d').strftime("%d.%m.%Y"),
        "contract_number":ws[f"B{k}"].value,
        "contract_date":'14.03.2023',#ws[f"L{k}"].value.replace(',', '.') if ws[f"L{k}"].value != None else '.'.join([klllll[:2], klllll[2:4], klllll[4:]])
        "total":round(isum, 2),
        "total_str":get_sum_str(isum, currency='usd'),
        "total_str_en":translate_text(get_sum_str(isum, currency='usd'), 'en'),
        "products":products,
    }
    doc = DocxTemplate("APP.docx")
    inner = 'done/'+str(datetime.now().timestamp()).split('.')[0]+f'APP {ws[f"C{k}"].value} {ws[f"B{k}"].value}.docx'.replace("/", "")
    doc.render(data)
    doc.save(inner)
    pppl += 1

    isum = 0
    x = 1

