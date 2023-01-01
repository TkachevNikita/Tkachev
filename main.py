import csv
import math
import matplotlib.pyplot as plt
import numpy as np
import doctest

class DataSet:
    ''' Класс для парсинга csv и создания словарей.

    Attributes:
        file_name (str): название файла
        profession_name (str): профессия, для которой производится выборка

    >>> type(DataSet('vacancies_by_year.csv', 'Аналитик')).__name__
    'DataSet'
    >>> DataSet('vacancies_by_year.csv', 'Аналитик').file_name
    'vacancies_by_year.csv'
    >>> DataSet('vacancies_by_year.csv', 'Аналитик').profession_name
    'Аналитик'
    >>> DataSet('vacancies_by_year.csv', 'Тестировщик').profession_name
    'Тестировщик'
    '''
    def __init__(self, file_name, profession_name):
        ''' Инициализирует объект DataSet.
        Args:
            file_name (str): название файла
            profession_name (str): профессия, для которой производится выборка
        '''
        self.file_name = file_name
        self.profession_name = profession_name

    def parse_csv(self):
        ''' Парсит csv и создает словари для дальнейшей работы

        :returns:
            dict: распределение средней зарплаты по годам
            dict: распределение кол-ва вакансий по годам
            dict: распределение средней зарплаты по годам для выбранной профессии
            dict: распределение кол-ва вакансий по годам для выбранной профессии
            dict: распределение средней зарплаты по городам, топ 10
            dict: распределение доли вакансий по городам (в процентах), топ 10
        '''
        with open(self.file_name, encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            titles, data = [], []
            isFirst = True
            for row in reader:
                if isFirst:
                    titles = row
                    isFirst = False
                    continue
                if len(titles) == len(row) and '' not in row:
                    dct = dict(zip(titles, row))
                    vacancy = Vacancy(dct)
                    data.append(vacancy)
            dct_years_count = {}
            dct_years_salary = {}
            dct_years_salary_filt = {}
            dct_years_count_filt = {}
            dct_salary_by_sity = {}
            dct_count_by_sity = {}
            for d in data:
                dct_salary_by_sity[d.area_name] = 0
                dct_count_by_sity[d.area_name] = 0
                dct_years_salary[int(d.published_at[:4])] = 0
                dct_years_count[int(d.published_at[:4])] = 0
                dct_years_salary_filt[int(d.published_at[:4])] = 0
                dct_years_count_filt[int(d.published_at[:4])] = 0
            for d in data:
                dct_count_by_sity[d.area_name] += 1
                dct_years_count[int(d.published_at[:4])] += 1
                dct_years_salary[int(d.published_at[:4])] += d.get_average()
                dct_salary_by_sity[d.area_name] += d.get_average()
                if self.profession_name in d.name:
                    dct_years_count_filt[int(d.published_at[:4])] += 1
                    dct_years_salary_filt[int(d.published_at[:4])] += d.get_average()
            delete = []
            for key in dct_count_by_sity:
                if round(dct_count_by_sity[key] / len(data), 4) < 0.01:
                    delete.append(key)
            dct_part = {}
            for i in delete:
                del dct_salary_by_sity[i]
            for key in dct_salary_by_sity:
                dct_salary_by_sity[key] = math.floor(dct_salary_by_sity[key] / dct_count_by_sity[key])
                dct_part[key] = round(dct_count_by_sity[key] / len(data), 4)
            for i in dct_years_count.keys():
                if dct_years_count[i] != 0:
                    dct_years_salary[i] = math.floor(dct_years_salary[i] / dct_years_count[i])
                if dct_years_count_filt[i] != 0:
                    dct_years_salary_filt[i] = math.floor(dct_years_salary_filt[i] / dct_years_count_filt[i])
        dct_salary_by_sity = sorted(dct_salary_by_sity.items(), key=lambda x: -x[1])
        dct_part = sorted(dct_part.items(), key=lambda x: -x[1])
        return dct_years_salary, dct_years_count, dct_years_salary_filt, dct_years_count_filt, dct_salary_by_sity, dct_part


class Vacancy:
    ''' Класс для представления вакансий.
        Attributes:
            name (str): название вакансии
            salary (экземпляр класса Salary): данные о зарплате
            area_name (str): место публикации вакансии
            published_at (str): дата и время публикации

    >>> Vacancy({'name':'Программист', 'salary_from':10, 'salary_to':20.5, 'salary_currency':'RUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'}).name
    'Программист'
    >>> type(Vacancy({'name':'Программист', 'salary_from':10, 'salary_to':20.5, 'salary_currency':'RUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'})).__name__
    'Vacancy'
    >>> Vacancy({'name':'Программист', 'salary_from':10, 'salary_to':20.5, 'salary_currency':'RUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'}).area_name
    'Москва'
    >>> Vacancy({'name':'Программист', 'salary_from':10, 'salary_to':20.5, 'salary_currency':'RUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'}).published_at
    '2007-12-03T17:34:36+0300'
    '''
    def __init__(self, dct):
        '''Инициализирует объект Vacancy.
            Args:
                dct (dict): словарь содержащий ключи из 1 сторки файла и значения - одна конкретная вакансия
        '''
        self.name = dct['name']
        self.salary = Salary(dct['salary_from'], dct['salary_to'], dct['salary_currency'])
        self.area_name = dct['area_name']
        self.published_at = dct['published_at']

    def get_average(self):
        ''' Вычисляет среднюю зарплату в рублях (конвертация с помоью словаря currency_to_rub)
        :returns:
            float: средняя зарплата в рублях

        >>> Vacancy({'name':'Программист', 'salary_from':10, 'salary_to':20.5, 'salary_currency':'RUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'}).get_average()
        15.0
        >>> Vacancy({'name':'Программист', 'salary_from':'10', 'salary_to':20.9999999, 'salary_currency':'RUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'}).get_average()
        15.0
        >>> Vacancy({'name':'Программист', 'salary_from':10, 'salary_to':'20', 'salary_currency':'EUR', 'area_name':'Москва', 'published_at':'2007-12-03T17:34:36+0300'}).get_average()
        898.5
        '''
        return 0.5 * (self.salary.salary_from * self.salary.currency_to_rub[self.salary.salary_currency] +
                      self.salary.salary_to * self.salary.currency_to_rub[self.salary.salary_currency])


class Salary:
    '''Класс для представления зарплаты.
    Attributes:
        salary_from (int): нижняя граница вилки оклада
        salary_to (int): верхняя граница вилки оклада
        salary_currency (str): валюта оклада

    >>> type(Salary(10.0, 20.4, 'RUR')).__name__
    'Salary'
    >>> Salary(10.0, 20.4, 'RUR').salary_to
    20
    >>> Salary(10.284, 20.4, 'RUR').salary_from
    10
    >>> Salary(10.0, 20.4, 'RUR').salary_currency
    'RUR'
    '''
    currency_to_rub = {
        "AZN": 35.68,
        "BYR": 23.91,
        "EUR": 59.90,
        "GEL": 21.74,
        "KGS": 0.76,
        "KZT": 0.13,
        "RUR": 1,
        "UAH": 1.64,
        "USD": 60.66,
        "UZS": 0.0055,
    }

    def __init__(self, salary_from, salary_to, salary_currency):
        ''' Инициализирует объект Salary.
        Args:
         salary_from (str or int or float): нижняя граница вилки оклада
         salary_to (str or int or float): верхняя граница вилки оклада
         salary_currency (str): валюта оклада
        '''
        self.salary_from = int(float(salary_from))
        self.salary_to = int(float(salary_to))
        self.salary_currency = salary_currency


class Report:
    '''Класс для формирования отчета.
        Attributes:
            dct_years_salary (dict): распределение уровня  средних зарплат по годам
            dct_years_count (dict): распределение количества вакансий по годам
            dct_years_salary_filt (dict): распределение уровня  средних зарплат по годам для выбранной профессии
            dct_years_count_filt (dict): распределение количества вакансий по годам для выбранной профессии
            dct_salary_by_sity (dict): распределение уровня средних зарплат по городам
            dct_part (dict): доля вакансий по городам (в процентах)

    >>> Report({'2007': 10000, '2020': 50000}, {}, {}, {}, [], []).dct_years_salary['2020']
    50000
    >>> Report({}, {'2007': 100, '2020': 50}, {}, {}, [], []).dct_years_count['2007']
    100
    >>> dict(Report({}, {}, {}, {}, [], [('Москва', 0.563), ('Санкт-Петербург', 0.115)]).dct_part)['Санкт-Петербург']
    0.115
    '''
    def __init__(self, dct_years_salary, dct_years_count, dct_years_salary_filt, dct_years_count_filt,
                 dct_salary_by_sity, dct_part):
        '''Инициализирует объект Report.
            Args:
                dct_years_salary (dict): распределение уровня  средних зарплат по годам
                dct_years_count (dict): распределение количества вакансий по годам
                dct_years_salary_filt (dict): распределение уровня  средних зарплат по годам для выбранной профессии
                dct_years_count_filt (dict): распределение количества вакансий по годам для выбранной профессии
                dct_salary_by_sity (list): распределение уровня средних зарплат по городам
                dct_part (list): доля вакансий по городам (в процентах)
        '''
        self.dct_years_salary = dct_years_salary
        self.dct_years_count = dct_years_count
        self.dct_years_salary_filt = dct_years_salary_filt
        self.dct_years_count_filt = dct_years_count_filt
        self.dct_salary_by_sity = dct_salary_by_sity
        self.dct_part = dct_part

    def generate_excel(self):
        '''Формирует отчет в виде таблицы.

        :returns:
            file (.xlsx): 'rep.xlsx'
        '''
        wb = openpyxl.Workbook()
        font = openpyxl.styles.Font(bold=True)
        thin = openpyxl.styles.Side(border_style="thin", color="000000")
        list1 = wb.active
        list1.title = "Статистика по годам"
        list2 = wb.create_sheet("Статистика по городам")
        list1.append(('Год', 'Средняя зарплата', f'Средняя зарплата - {profession_name}', 'Количество вакансий',
                      f'Количество вакансий - {profession_name}'))

        list1.column_dimensions['A'].width = 6
        list1.column_dimensions['B'].width = len('Средняя зарплата') + 1
        list1.column_dimensions['C'].width = len(f'Средняя зарплата - {profession_name}') + 1
        list1.column_dimensions['D'].width = len('Количество вакансий') + 1
        list1.column_dimensions['E'].width = len(f'Количество вакансий - {profession_name}') + 1

        list1['A1'].font = font
        list1['B1'].font = font
        list1['C1'].font = font
        list1['D1'].font = font
        list1['E1'].font = font

        for year in self.dct_years_salary.keys():
            list1.append((year, self.dct_years_salary[year], self.dct_years_salary_filt[year],
                          self.dct_years_count[year], self.dct_years_count_filt[year]))

        for i in range(1, len(self.dct_years_salary) + 2):
            a = f'A{i}'
            b = f'B{i}'
            c = f'C{i}'
            d = f'D{i}'
            e = f'E{i}'
            list1[a].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list1[b].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list1[c].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list1[d].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list1[e].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)

        list2.append(('Город', 'Уровень зарплат', ' ', 'Город', 'Доля вакансий'))

        list2.column_dimensions['A'].width = 1
        list2.column_dimensions['B'].width = 17
        list2.column_dimensions['C'].width = 2
        list2.column_dimensions['D'].width = 1
        list2.column_dimensions['E'].width = 15

        list2['A1'].font = font
        list2['B1'].font = font
        list2['D1'].font = font
        list2['E1'].font = font

        i = 2
        for city in dict(self.dct_salary_by_sity).keys():
            if i < 12:
                list2.append((city, dict(self.dct_salary_by_sity)[city]))
                if len(city) > list2.column_dimensions['A'].width:
                    list2.column_dimensions['A'].width = len(city) + 2
            i += 1

        j = 2
        for city in dict(self.dct_part).keys():
            if j < 12:
                d = f'D{j}'
                e = f'E{j}'
                list2[d].value = city
                list2[e].value = f'{round(dict(self.dct_part)[city]*100, 2)}%'
                list2[e].alignment = openpyxl.styles.Alignment(horizontal='right')
                if len(city) > list2.column_dimensions['D'].width:
                    list2.column_dimensions['D'].width = len(city) + 2
            j += 1

        for i in range(1, 12):
            a = f'A{i}'
            b = f'B{i}'
            d = f'D{i}'
            e = f'E{i}'

            list2[a].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list2[b].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list2[d].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)
            list2[e].border = openpyxl.styles.Border(top=thin, left=thin, bottom=thin, right=thin)

        wb.save('rep.xlsx')

    def generate_image(self):
        ''' Формирует отчет в виде графиков (изображение).

        :returns:
            file (.png): 'graph.png'
        '''
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(nrows=2, ncols=2)
        index = np.arange(len(self.dct_years_salary))
        index2 = np.arange(10)
        index_years = []
        values_all = []
        count_all = []
        city_ax3 = []
        city_salary = []
        count_by_prof = []
        values_by_prof = []
        city_parts = []
        num_parts = []
        for key in dict(self.dct_part[:10]).keys():
            city_parts.append(key)
            num_parts.append(dict(self.dct_part)[key])
        city_parts.append('Другие')
        num_other = 1 - sum(dict(self.dct_part[:10]).values())
        num_parts.append(num_other)
        for key in dict(self.dct_salary_by_sity[:10]).keys():
            city_ax3.append(key.replace(" ", "\n").replace("-", "-\n"))
            city_salary.append(dict(self.dct_salary_by_sity)[key])
        for value in self.dct_years_count.values():
            count_all.append(value)
        for value in self.dct_years_count_filt.values():
            count_by_prof.append(value)
        for value in self.dct_years_salary_filt.values():
            values_by_prof.append(value)
        for key in self.dct_years_salary.keys():
            index_years.append(key)
            values_all.append(self.dct_years_salary[key])

        ax1.set_title('Уровень зарплат по годам', fontsize=8)
        ax1.bar(index, values_all, 0.4, label='Средняя з/п')
        ax1.bar(index+0.4, values_by_prof, 0.4, label=f'з/п {profession_name.lower()}')
        ax1.set_xticks(index + 0.2, index_years, rotation=90)
        ax1.legend(loc=2)
        ax1.xaxis.set_tick_params(labelsize=8)
        ax1.yaxis.set_tick_params(labelsize=8)
        ax1.yaxis.grid(True)

        ax2.set_title('Количество вакансий по годам', fontsize=8)
        ax2.bar(index, count_all, 0.4, label='Количество вакансий')
        ax2.bar(index + 0.4, count_by_prof, 0.4, label=f'Количество вакансий {profession_name.lower()}')
        ax2.set_xticks(index + 0.2, index_years, rotation=90)
        ax2.legend(loc=1)
        ax2.xaxis.set_tick_params(labelsize=8)
        ax2.yaxis.set_tick_params(labelsize=8)
        ax2.yaxis.grid(True)

        ax3.set_title('Уровень зарплат по городам', fontsize=8)
        ax3.barh(index2, list(reversed(city_salary)), 0.6)
        ax3.set_yticks(index2, list(reversed(city_ax3)))
        ax3.xaxis.set_tick_params(labelsize=8)
        ax3.yaxis.set_tick_params(labelsize=6)
        ax3.xaxis.grid(True)

        ax4.set_title('Доля вакансий по городам', fontsize=8)
        ax4.pie(list(reversed(num_parts)), labels=list(reversed(city_parts)), textprops={'fontsize': 6})
        ax4.axis('equal')

        plt.tight_layout()
        plt.savefig('graph.png')



print_with_input = 'Введите данные для печати: '
user_waiting = input('Требуемый формат вывода (Вакансии или Статистика): ')
file_name, profession_name = input(print_with_input).split()
dataset = DataSet(file_name, profession_name)
dct_years_salary, dct_years_count, dct_years_salary_filt, dct_years_count_filt, dct_salary_by_sity, dct_part = dataset.parse_csv()

report = Report(dct_years_salary, dct_years_count, dct_years_salary_filt, dct_years_count_filt, dct_salary_by_sity,
                dct_part)

if user_waiting.lower() == 'вакансии':
    report.generate_image()
elif user_waiting.lower() == 'статистика':
    report.generate_excel()
else:
    print("Некорректный формат вывода")

