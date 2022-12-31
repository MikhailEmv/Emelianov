from csv import reader, writer


def parse_csv_file(file="C:/Users/Michael/PycharmProjects/Emelianov/vacancies_big.csv") -> tuple:
    """Функция фильтрации файла от некорректных строк и создания словаря с вакансиями по годам

    :param file: директория .csv файла
    :return: кортеж словаря с вакансиями по годам и headline листа
    """
    years_dictionary = dict()
    with open(file, 'r', encoding='utf-8-sig') as current_file:
        csv_reader = reader(current_file)
        headline = next(csv_reader)
        headline[0] = 'name'

        for item in list(csv_reader):
            if len(item) == len(headline) and '' not in item:
                current_year = item[headline.index('published_at;;;')][:4]
                if current_year not in years_dictionary.keys():
                    years_dictionary[current_year] = [item]
                else:
                    years_list = years_dictionary[current_year]
                    years_list.append(item)
                    years_dictionary[current_year] = years_list
    return years_dictionary, headline


def create_csv_files(headlines: list, years_vacancies: list) -> None:
    """Создание новых csv файлов в текущей папке

    :param headlines: список заголовков
    :param years_vacancies: словарь со списками вакансий
    :return: nothing
    """
    for current_year in years_vacancies:
        with open(f'CSV/{current_year}.csv', 'w', encoding='utf-8-sig') as following_file:
            csv_writer = writer(following_file)
            csv_writer.writerow(headlines)
            for current_row in years_vacancies[current_year]:
                csv_writer.writerow(current_row)


csv_dictionary, headline = parse_csv_file()
create_csv_files(headline, csv_dictionary)
