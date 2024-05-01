import csv
import shutil

import openpyxl
import pandas as pd
import re

import pandas
import xlsxwriter


def remove_difference(some_str):
    if 'Навчальна практика (за фахом)' in str(some_str):
        some_str = 'Навчальні практика (за фахом)'
    new_str = (str(some_str).replace('A', 'А')
               .replace('проектов', 'проєктів')
               .replace('проектами', 'проєктами')
               .replace('в тому числі', 'у тому числі')
               .replace('Технологія розробки та експлуатації інформаційних систем військового призначення',
                        'Технології розробки інформаційних систем військового призначення')
               .replace('Військова педагогіка та психологія (у тому числі лідерство)',
                        'Військова педагогіка та психологія')
               .replace('Бойове застосування безпілотних авіаційних комплексів ретрансляторів',
                        'Бойове застосування безпілотних авіаційних комплексів')
               .replace('Комплексний екзамен з фахової підготовки', 'Комплексний екзамен зі спеціальності')
               .replace('програмниого', 'програмного')
               .replace('B', 'В')
               .replace('C', 'С')
               .replace('E', 'Е')
               .replace('H', 'Н')
               .replace('I', 'І')
               .replace('K', 'К')
               .replace('M', 'М')
               .replace('O', 'О')
               .replace('P', 'Р')
               .replace('T', 'Т')
               .replace('X', 'Х')
               .replace('a', 'а')
               .replace('c', 'с')
               .replace('e', 'е')
               .replace('i', 'і')
               .replace('k', 'к')
               .replace('o', 'о')
               .replace('p', 'р')
               .replace('x', 'х')
               .replace('y', 'у')
               .replace('`', "'")
               .replace('’', "'")
               .split(' '))
    pattern = r'\s+'
    for sub_str in new_str:

        new_str[new_str.index(sub_str)] = re.sub(pattern, '', sub_str)
        if sub_str == '':
            new_str.remove(sub_str)
    return ' '.join(new_str).strip(' ')


def process(setup):
    # get lessons from template sheet
    excel_lessons = pandas.read_excel(setup['template'])
    lessons_template = []
    arr_len = len(excel_lessons)
    i = 0
    while i < arr_len:
        if (type(excel_lessons.iloc[i].values.tolist()[1]) is not type(0.0) and
                type(excel_lessons.iloc[i].values.tolist()[0]) is type(1)):
            if excel_lessons.iloc[i].values.tolist()[1].split('/') and int(
                    excel_lessons.iloc[i].values.tolist()[0]) > 0:
                lessons_template.append(
                    [remove_difference(
                        ' '.join(excel_lessons.iloc[i].values.tolist()[1].split('/')[0].split(' ')).strip(' ')), i + 2])
        i += 1
    print(lessons_template)
    print(len(lessons_template))

    print('\n')
    excel_reader = pd.ExcelFile(setup['source'])
    excel = excel_reader.parse(setup['sheet'])  # .values.tolist()

    pattern_group = fr"{setup['group']}\s*нг"
    print(pattern_group)
    pattern_PIB = r"^[А-ЩЬЮЯЇІЄҐ][а-щьюяїієґ]+\s*[А-ЩЬЮЯЇІЄҐ]\.(\s*)+[А-ЩЬЮЯЇІЄҐ]\.$"
    arr_len = len(excel)
    print(arr_len)
    lessons_diploma = []

    i = 0
    while i < arr_len:
        row = excel.iloc[i].values.tolist()
        if re.fullmatch(pattern_group, str(row[2])):
            for item in row:
                row[row.index(item)] = remove_difference(item)

            for item in row:
                if type(item) is type('str'):
                    split_temp = remove_difference(item)
                    for lesson in lessons_template:

                        if re.sub(r'\s', '', split_temp.lower()) == re.sub(r'\s', '', lesson[0].lower()):
                            if re.search(r'^\d+$',
                                         remove_difference(str(excel.iloc[i + 2].values.tolist()[row.index(item)]))):
                                lessons_diploma.append(
                                    {'lesson': split_temp,
                                     'source_position': row.index(item),
                                     'credits': excel.iloc[i + 1].values.tolist()[row.index(item)],
                                     'template_position': lesson[1]})

                                row[row.index(item) + 1] = ['meow']
                                row[row.index(item) + 2] = ['meow']
                                row[row.index(item)] = ['meow']
                                lessons_template[lessons_template.index(lesson)] = 'meow meow'
                                break

                        if split_temp == 'Комплексний екзамен з фахової підготовки':
                            lessons_diploma.append(
                                {'lesson': split_temp,
                                 'source_position': row.index('Комплексний екзамен з фахової підготовки'),
                                 'credits': excel.iloc[i + 1].values.tolist()[
                                     row.index('Комплексний екзамен з фахової підготовки')],
                                 'template_position': lesson[1]})
                            print(lesson)
                            row[row.index(item) + 1] = ['meow']
                            row[row.index(item) + 2] = ['meow']
                            row[row.index(item)] = ['meow']
                            break
            for someme in lessons_diploma:
                print(someme)
            print(len(lessons_diploma))
            print(lessons_template)
            print(row)
            print('\n----------------------------------------------------\ncadet')

        if re.fullmatch(pattern_PIB, str(row[2])) and row[0] == int(setup['group']):
            print(excel.iloc[i].values.tolist()[2])
            score_arr_unfiltered = []
            for item in row:
                if len(str(item)) == 2:
                    score_arr_unfiltered.append([item, row.index(item)])
                if item in ['Неприйнятно', 'Незадовільно', 'Достатньо', 'Задовільно', 'Добре', 'Дуже добре',
                            'Відмінно']:
                    score_arr_unfiltered.append([item, row.index(item)])
                row[row.index(item)] = ['olololo']
            score_arr_filtered = []
            j = 0
            while j < len(score_arr_unfiltered):
                if j + 1 != len(score_arr_unfiltered):
                    if type(score_arr_unfiltered[j][0]) == type(0) and type(score_arr_unfiltered[j + 1][0]) == type(
                            'str'):
                        score_arr_filtered.append({'score': score_arr_unfiltered[j][0],
                                                   'source_position': score_arr_unfiltered[j][1]})
                j += 1
            print(score_arr_filtered)

            file_name = re.sub(r'\s+', '', excel.iloc[i].values.tolist()[2])
            file_name = re.sub(r'\.', '', file_name)
            copy_workbook = xlsxwriter.Workbook(f"{setup['dist']}/{file_name}.xlsx")
            copy_worksheet = copy_workbook.add_worksheet('Table')
            copy_workbook.close()

            shutil.copyfile(setup['template'],
                            f"{setup['dist']}/{file_name}.xlsx")

            workbook = openpyxl.load_workbook(setup['template'])
            worksheet = workbook['Table']

            for lesson in lessons_diploma:
                print(lesson)
                for score in score_arr_filtered:
                    print(score)
                    if lesson['source_position'] == score['source_position']:
                        worksheet[f"L{lesson['template_position']}"] = lesson['credits']
                        worksheet[f"N{lesson['template_position']}"] = score['score']
            workbook.save(f"{setup['dist']}/{file_name}.xlsx")
        i += 1
