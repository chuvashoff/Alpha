import openpyxl
import os

import datetime
from my_func import *
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


path_config = input('Укажите путь до конфигуратора\n')
print(datetime.datetime.now())

file_config = 'UnimodCreate.xlsm'

'''Ищем файл конфигуратора в указанном каталоге'''
for file in os.listdir(path_config):
    if file.endswith('.xlsm') or file.endswith('.xls'):
        file_config = file
        break


all_CPU = []
pref_IP = []
sl_object_rus = {}
sl_object_all = {}
num_pz = 0

'''Считываем файл-шаблон для AI  AE SET'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_AIAESET'), 'r', encoding='UTF-8') as f:
    tmp_object_AIAESET = f.read()
'''Считываем файл-шаблон для DI'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_DI'), 'r', encoding='UTF-8') as f:
    tmp_object_DI = f.read()
'''Считываем файл-шаблон для IM'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_IM'), 'r', encoding='UTF-8') as f:
    tmp_object_IM = f.read()
'''Считываем файл-шаблон для BTN CNT'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_BTN_CNT'), 'r', encoding='UTF-8') as f:
    tmp_object_BTN_CNT = f.read()
'''Считываем файл-шаблон для PZ'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_PZ'), 'r', encoding='UTF-8') as f:
    tmp_object_PZ = f.read()
'''Считываем файл-шаблон для группы'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_group'), 'r', encoding='UTF-8') as f:
    tmp_group = f.read()
'''Считываем файл-шаблон для app'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_app'), 'r', encoding='UTF-8') as f:
    tmp_app = f.read()
'''Считываем файл-шаблон для контроллера'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_TREI'), 'r', encoding='UTF-8') as f:
    tmp_trei = f.read()
'''Считываем файл-шаблон для global'''
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_global'), 'r', encoding='UTF-8') as f:
    tmp_global = f.read()

book = openpyxl.open(os.path.join(path_config, file_config), read_only=True)
'''читаем список всех контроллеров'''
sheet = book['Настройки']  # worksheets[1]
cells = sheet['B2': 'B22']
for p in cells:
    if p[0].value is not None:
        all_CPU.append(p[0].value)
'''Читаем префиксы IP адреса ПЛК'''
cells = sheet['B45': 'B46']
for p in cells:
    pref_IP.append(p[0].value + '.')

'''Читаем состав объектов'''
cells = sheet['B24': 'R38']
for p in cells:
    if p[0].value is not None:
        sl_object_rus[p[0].value] = p[1].value
        tmp0 = []
        tmp1 = []
        for i in range(12, 17):
            if p[i].value == 'ON':
                tmp0.append(p[i - 10].value)
                tmp1.append(p[i - 5].value)
        sl_object_all[p[0].value] = [tmp0, tmp1]

ff = open('file_plc.txt', 'w', encoding='UTF-8')
ff.close()
'''Далее для всех контроллеров, что нашли, делаем'''
for i in all_CPU:
    '''Измеряемые'''
    sheet = book['Измеряемые']  # .worksheets[3]
    cells = sheet['A2': 'AE' + str(sheet.max_row + 1)]
    sl_CPU_one = is_load_ai_ae_set(i, cells)

    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_ai_ae_set(sl_CPU_one, tmp_object_AIAESET, 'Types.AI.PLC_AI')

        with open('file_out_group.txt', 'w', encoding='UTF-8') as f:
            f.write(Template(tmp_group).substitute(name_group='AI', objects=tmp_line_))

    '''Расчетные'''
    sheet = book['Расчетные']  # .worksheets[4]
    cells = sheet['A2': 'AE' + str(sheet.max_row)]
    sl_CPU_one = is_load_ai_ae_set(i, cells)

    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_ai_ae_set(sl_CPU_one, tmp_object_AIAESET, 'Types.AE.PLC_AE')

        with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
            f.write(Template(tmp_group).substitute(name_group='AE', objects=tmp_line_))

    '''Дискретные'''
    sheet = book['Входные']  # .worksheets[6]
    cells = sheet['A2': 'W' + str(sheet.max_row)]
    sl_CPU_one = is_load_di(i, cells)

    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_di(sl_CPU_one, tmp_object_DI, 'Types.DI.PLC_DI')

        with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
            f.write(Template(tmp_group).substitute(name_group='DI', objects=tmp_line_))

    '''ИМ'''

    sheet = book['ИМ']  # .worksheets[9]
    cells = sheet['A2': 'T' + str(sheet.max_row)]
    sl_CPU_one = is_load_im(i, cells)
    sl_cnt = {}
    for key, value in sl_CPU_one.items():
        if value[4] == 'Да':
            sl_cnt[key+'_Worktime'] = [value[0]]
        if value[5] == 'Да':
            sl_cnt[key+'_Swap'] = [value[0]]

    '''ИМ АО- объединяем словари с ИМами'''
    sheet = book['ИМ(АО)']  # .worksheets[8]
    cells = sheet['A2': 'AA' + str(sheet.max_row)]
    sl_CPU_one = {**sl_CPU_one, **is_load_im_ao(i, cells)}

    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_im(sl_CPU_one, tmp_object_IM)

        with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
            f.write(Template(tmp_group).substitute(name_group='IM', objects=tmp_line_))

    '''Кнопки(в составе System)'''
    sheet = book['Кнопки']  # .worksheets[10]
    cells = sheet['A2': 'C' + str(sheet.max_row)]
    sl_CPU_one = is_load_btn(i, cells)

    tmp_subgroup = ''
    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_btn_cnt(sl_CPU_one, tmp_object_BTN_CNT, 'Types.BTN.PLC_BTN')
        tmp_subgroup += Template(tmp_group).substitute(name_group='BTN', objects=tmp_line_)

    '''Уставки(в составе System)'''
    sheet = book['Уставки']  # .worksheets[5]
    cells = sheet['A2': 'AG' + str(sheet.max_row)]
    sl_CPU_one = is_load_ai_ae_set(i, cells)

    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_ai_ae_set(sl_CPU_one, tmp_object_AIAESET, 'Types.SET.PLC_SET')
        tmp_subgroup += Template(tmp_group).substitute(name_group='SET', objects=tmp_line_)

    '''Наработки(CNT- прочитали при ИМ) в составе System)'''
    if len(sl_cnt) != 0:
        tmp_line_ = is_create_objects_btn_cnt(sl_cnt, tmp_object_BTN_CNT, 'Types.CNT.PLC_CNT')
        tmp_subgroup += Template(tmp_group).substitute(name_group='CNT', objects=tmp_line_)

    '''Защиты(PZ) в составе System'''
    sheet = book['Сигналы']  # .worksheets[11]
    cells = sheet['A2': 'N' + str(sheet.max_row)]
    sl_CPU_one, num_pz = is_load_pz(i, cells, num_pz)

    if len(sl_CPU_one) != 0:
        tmp_line_ = is_create_objects_pz(sl_CPU_one, tmp_object_PZ, 'Types.CNT.PLC_CNT')
        tmp_subgroup += Template(tmp_group).substitute(name_group='PZ', objects=tmp_line_)

    '''Формируем подгруппу'''
    if tmp_subgroup != '':
        with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
            f.write(Template(tmp_group).substitute(name_group='System', objects=tmp_subgroup.rstrip()))

    '''Формирование выходного файла app'''
    with open('file_out_group.txt', 'r', encoding='UTF-8') as f:
        tmp_line_ = f.read().rstrip()
    with open('file_app_out.txt', 'w', encoding='UTF-8') as f:
        f.write(Template(tmp_app).substitute(name_app='Tree', ct_object=tmp_line_))

    with open('file_app_out.txt', 'r', encoding='UTF-8') as f:
        tmp_line_ = f.read().rstrip()

    '''Для каждого объекта создаём контроллер '''
    num_obj_plc = 1
    for obj in sl_object_all:
        if i in sl_object_all[obj][0]:
            index_tmp = sl_object_all[obj][0].index(i)
            with open('file_plc.txt', 'a', encoding='UTF-8') as f:
                f.write(Template(tmp_trei).substitute(plc_name='PLC_' + i + str(num_obj_plc), plc_name_type='m903e',
                                                      ip_eth1=pref_IP[0] + sl_object_all[obj][1][index_tmp],
                                                      ip_eth2=pref_IP[1] + sl_object_all[obj][1][index_tmp],
                                                      dp_app=tmp_line_))
            num_obj_plc += 1
        else:
            num_obj_plc += 1
            continue

    '''чистим файл групп для корректной обработки отсуствия'''
    file_clear = open('file_out_group.txt', 'w', encoding='UTF-8')
    file_clear.close()

'''Формирование выходного файла для каждого контроллера'''
with open('file_plc.txt', 'r', encoding='UTF-8') as f:
    tmp_line_ = f.read().rstrip()

with open('file_out_plc_aspect.omx-export', 'w', encoding='UTF-8') as f:
    f.write(Template(tmp_global).substitute(dp_node=tmp_line_))


book.close()
print(datetime.datetime.now())


os.remove('file_plc.txt')
os.remove('file_out_group.txt')
'''os.remove('file_out_objects.txt')'''
os.remove('file_app_out.txt')
