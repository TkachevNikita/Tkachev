import csv
from datetime import datetime
import datefinder
import cProfile

data = []
with open('vacancies_by_year.csv', encoding='utf-8-sig') as file:
    data = [row[5] for row in csv.reader(file) if 'published_at' not in row]


def profile(func):
    def wrapper(data_list):
        profile = cProfile.Profile()
        profile.enable()
        f = [func(d) for d in data_list]
        profile.disable()
        print(f'{func.__name__}:')
        profile.print_stats(0)

    return wrapper


def datetime_test(date):
    date = datetime.strptime(date[:10], '%Y-%m-%d').date()
    return f'{date.day}.{date.month}.{date.year}'


def slice_test(date):
    day = date[8:10]
    month = date[5:7]
    year = date[:4]
    return f'{day}.{month}.{year}'


def split_test(date):
    date = date.split('T')[0].split('-')
    day = date[2]
    month = date[1]
    year = date[0]
    return f'{day}.{month}.{year}'


def datefinder_test(date):
    matches = list(datefinder.find_dates(date))
    date = str(matches[0])
    return f'{date[8:10]}.{date[5:7]}.{date[:4]}'


pr = profile(split_test)
pr(data)
