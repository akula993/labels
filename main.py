import os
from pathlib import Path

from docxtpl import DocxTemplate
from docx2pdf import convert
import PySimpleGUI as sg


def ev(file_result):
    if values['party']:
        party = values['party']
    else:
        party = str('')
    if values['net_weight']:
        net_weight = values['net_weight']
    else:
        net_weight = str('')
    if values['date_of_shipment']:
        date_of_shipment = values['date_of_shipment']
    else:
        date_of_shipment = str('')

    if values['number']:
        number = values['number']
    else:
        number = ''
    file_one = './шаблон2.docx'
    doc = DocxTemplate(file_one)
    context = {'charge': party, 'nettogewicht': net_weight, 'versanddatum': date_of_shipment,
               'number': number, }
    file_result = values['file_result']

    doc.render(context)
    fol = values['file']
    os.chdir(fol)
    doc.save(file_result + ".docx")
    convert(file_result + ".docx", file_result + ".pdf")

my_home = os.chdir(f'{Path.home()}')

sg.theme('Dark Black')

layout = [
    [sg.Text('Партия', size=(15, 1)), sg.InputText(key='party', default_text='0000370533')],
    [sg.Text('Вес нетто', size=(15, 1)), sg.InputText(key='net_weight', default_text='30kg')],
    [sg.Text('Дата отгрузки', size=(15, 1)), sg.InputText(key='date_of_shipment', default_text='13.10.2022')],
    [sg.Text('Номер', size=(15, 1)), sg.InputText(key='number', default_text='2622 VS')],
    [sg.Text('Название файла', size=(15, 1)), sg.InputText(default_text='test', key='file_result')],
    [sg.Text('Выбор папки обязателен*')],
    [sg.FolderBrowse('Папка куда сохранять', key='file', enable_events=False), sg.Text()],
    [sg.Submit("Создание"), sg.Cancel("Выход")]
]

window = sg.Window('Создание файлов этикетки', layout)

while True:
    event, values = window.read()

    if event in (None, 'Exit', 'Cancel', "Выход",):
        break

    if event == 'Создание':
        ev(values['file_result'])

window.close()
