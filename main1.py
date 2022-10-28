from docxtpl import DocxTemplate
from docx2pdf import convert
import PySimpleGUI as sg

sg.theme('Dark Red')

layout = [
    [sg.FileBrowse('Файл', size=(15, 1), key='file1'), sg.Text('Выбраный файл:')],
    # [sg.Text(f'Выбраный файл:{open("шаблон1.docx", "r")}', key='file1',), sg.Input(key='-FILE-', visible=False, enable_events=True)],
    # [sg.Text('Title', size=(15, 1)), sg.InputText(key='title',
    #                                               default_text='LLC “Grafo Impex” ulitsa Materkova, dom 4, etage 2, pomeshenie 1, office 102 115280 Moscow RUSSISCHE FODERATION')],
    [sg.Text('Партия', size=(15, 1)), sg.InputText(key='party', default_text='0000370533')],
    [sg.Text('Вес нетто', size=(15, 1)), sg.InputText(key='net_weight', default_text='30kg')],
    [sg.Text('Дата отгрузки', size=(15, 1)), sg.InputText(key='date_of_shipment', default_text='13.10.2022')],
    [sg.Text('Номер', size=(15, 1)), sg.InputText(key='number', default_text='2622 VS')],

    [sg.Text('Название файла', size=(15, 1)), sg.InputText(default_text='test', key='file_result')],
    [sg.Submit("Создание"), sg.Cancel("Выход")]
]
window = sg.Window('Создание файлов этикетки', layout)
while True:  # The Event Loop
    event, values = window.read()
    file_old = values['file1']
    
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
    file_result = values['file_result']
    if event in (None, 'Exit', 'Cancel', "Выход",):
        break
    if event == 'Создание':
        doc = DocxTemplate(file_old)
        context = {'charge': party, 'nettogewicht': net_weight, 'versanddatum': date_of_shipment,
                   'number': number, }
        doc.render(context)
        doc.save(file_result + ".docx")
        convert(file_result + ".docx", file_result + ".pdf")

window.close()

# with open('1_file.json', 'w', encoding=encoding) as file:
#     json.dump(file_name, file, indent=4, ensure_ascii=False)
