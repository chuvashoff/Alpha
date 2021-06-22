import copy
import openpyxl
import datetime
import logging
from my_func import *
from alpha_index import *
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try:
    path_config = input('Укажите путь до конфигуратора\n')
    print(datetime.datetime.now())

    file_config = 'UnimodCreate.xlsm'

    '''Ищем файл конфигуратора в указанном каталоге'''
    for file in os.listdir(path_config):
        if file.endswith('.xlsm') or file.endswith('.xls'):
            file_config = file
            break

    all_CPU = []
    sl_CPU_spec = {}
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
    with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_BTN_CNT_sig'), 'r', encoding='UTF-8') as f:
        tmp_object_BTN_CNT_sig = f.read()  # его же используем для диагностики CPU потому что подходит
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
    '''Считываем файл-шаблон для топливного регулятора'''
    with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_TR_ps90'), 'r', encoding='UTF-8') as f:
        tmp_tr_ps90 = f.read()
    '''Считываем файл-шаблон для АПР'''
    with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_APR'), 'r', encoding='UTF-8') as f:
        tmp_apr = f.read()

    book = openpyxl.open(os.path.join(path_config, file_config))  # , read_only=True
    '''читаем список всех контроллеров'''
    sheet = book['Настройки']  # worksheets[1]
    cells = sheet['B2': 'B22']
    for p in cells:
        if p[0].value is not None:
            all_CPU.append(p[0].value)
    '''Читаем префиксы IP адреса ПЛК(нужно продумать про новые конфигураторы)'''
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
    '''Мониторинг ТР и АПР в контроллере'''
    cells = sheet['B1':'L21']
    for p in cells:
        if p[0].value is None:
            break
        else:
            sl_CPU_spec[p[0].value] = []
            if p[is_f_ind(cells[0], 'FLR')].value == 'ON' and p[is_f_ind(cells[0], 'Тип ТР')].value == 'ПС90':
                sl_CPU_spec[p[0].value].append('ТР')
            if p[is_f_ind(cells[0], 'APR')].value == 'ON':
                sl_CPU_spec[p[0].value].append('АПР')

    ff = open('file_plc.txt', 'w', encoding='UTF-8')
    ff.close()
    '''Далее для всех контроллеров, что нашли, делаем'''
    for i in all_CPU:
        '''Диагностика модулей (DIAG)'''
        sheet = book['Модули']
        cells = sheet['A1': 'G' + str(sheet.max_row)]
        sl_modules_cpu = {}
        for p in cells:
            if p[is_f_ind(cells[0], 'CPU')].value == i:
                aa = copy.copy(sl_modules[p[is_f_ind(cells[0], 'Шифр модуля')].value])
                sl_modules_cpu[p[is_f_ind(cells[0], 'Имя модуля')].value] = [p[is_f_ind(cells[0], 'Шифр модуля')].value, aa]

        for jj in ['Измеряемые', 'Входные', 'Выходные', 'ИМ(АО)']:
            sheet_run = book[jj]
            cells_run = sheet_run['A1': 'O' + str(sheet_run.max_row)]
            for p in cells_run:
                if p[is_f_ind(cells_run[0], 'CPU')].value == i and p[is_f_ind(cells_run[0], 'Нестандартный канал')].value == 'Нет':
                    tmp_ind = int(p[is_f_ind(cells_run[0], 'Номер канала')].value) - 1
                    sl_modules_cpu[p[is_f_ind(cells_run[0], 'Номер модуля')].value][1][tmp_ind] = \
                        is_cor_chr(p[is_f_ind(cells_run[0], 'Наименование параметра')].value)

        if len(sl_modules_cpu) != 0:
            tmp_line_ = is_create_objects_diag(sl_modules_cpu, tmp_object_BTN_CNT_sig)
            tmp_line_ = (Template(tmp_group).substitute(name_group='HW', objects=tmp_line_))

            with open('file_out_group.txt', 'w', encoding='UTF-8') as f:
                f.write(Template(tmp_group).substitute(name_group='Diag', objects=tmp_line_.rstrip()))

        '''Измеряемые'''
        sheet = book['Измеряемые']  # .worksheets[3]
        cells = sheet['A1': 'AG' + str(sheet.max_row)]
        sl_CPU_one = is_load_ai_ae_set(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                       is_f_ind(cells[0], 'Наименование параметра'),
                                       is_f_ind(cells[0], 'Единицы измерения'),
                                       is_f_ind(cells[0], 'Короткое наименование'),
                                       is_f_ind(cells[0], 'Количество знаков'),
                                       is_f_ind(cells[0], 'CPU'))

        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_ai_ae_set(sl_CPU_one, tmp_object_AIAESET, 'Types.AI.AI_PLC_View')

            with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
                f.write(Template(tmp_group).substitute(name_group='AI', objects=tmp_line_))

        '''Расчетные'''
        sheet = book['Расчетные']  # .worksheets[4]
        cells = sheet['A1': 'AE' + str(sheet.max_row)]
        sl_CPU_one = is_load_ai_ae_set(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                       is_f_ind(cells[0], 'Наименование параметра'),
                                       is_f_ind(cells[0], 'Единицы измерения'),
                                       is_f_ind(cells[0], 'Короткое наименование'),
                                       is_f_ind(cells[0], 'Количество знаков'),
                                       is_f_ind(cells[0], 'CPU'))

        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_ai_ae_set(sl_CPU_one, tmp_object_AIAESET, 'Types.AE.AE_PLC_View')

            with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
                f.write(Template(tmp_group).substitute(name_group='AE', objects=tmp_line_))

        '''Дискретные'''
        sheet = book['Входные']  # .worksheets[6]
        cells = sheet['A1': 'AC' + str(sheet.max_row)]
        sl_CPU_one, sl_wrn = is_load_di(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                        is_f_ind(cells[0], 'ИМ'),
                                        is_f_ind(cells[0], 'Наименование параметра'),
                                        is_f_ind(cells[0], 'Цвет при наличии'),
                                        is_f_ind(cells[0], 'Цвет при отсутствии'),
                                        is_f_ind(cells[0], 'Предупреждение'),
                                        is_f_ind(cells[0], 'Текст предупреждения'),
                                        is_f_ind(cells[0], 'CPU'))

        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_di(sl_CPU_one, tmp_object_DI, 'Types.DI.DI_PLC_View')

            with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
                f.write(Template(tmp_group).substitute(name_group='DI', objects=tmp_line_))

        '''ИМ'''
        sheet = book['ИМ']  # .worksheets[9]
        cells = sheet['A1': 'T' + str(sheet.max_row)]
        sl_CPU_one = is_load_im(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                is_f_ind(cells[0], 'Наименование параметра'),
                                is_f_ind(cells[0], 'Тип ИМ'),
                                is_f_ind(cells[0], 'Род'),
                                is_f_ind(cells[0], 'Считать наработку'),
                                is_f_ind(cells[0], 'Считать перестановки'),
                                is_f_ind(cells[0], 'CPU'))
        sl_cnt = {}
        for key, value in sl_CPU_one.items():
            if value[4] == 'Да':
                sl_cnt[key+'_WorkTime'] = [value[0]]
            if value[5] == 'Да':
                sl_cnt[key+'_Swap'] = [value[0]]

        '''ИМ АО- объединяем словари с ИМами'''
        sheet = book['ИМ(АО)']  # .worksheets[8]
        cells = sheet['A1': 'AA' + str(sheet.max_row)]
        sl_CPU_one = {**sl_CPU_one, **is_load_im_ao(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                                    is_f_ind(cells[0], 'Наименование параметра'),
                                                    is_f_ind(cells[0], 'Род'),
                                                    is_f_ind(cells[0], 'ИМ'),
                                                    is_f_ind(cells[0], 'CPU'))}

        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_im(sl_CPU_one, tmp_object_IM)

            with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
                f.write(Template(tmp_group).substitute(name_group='IM', objects=tmp_line_))

        '''Добавляем АПР, если для данного контроллера указан АПР в настройках'''
        if 'АПР' in sl_CPU_spec[i]:
            with open('file_out_group.txt', 'a', encoding='UTF-8') as f:
                f.write(tmp_apr.rstrip())

        '''Кнопки(в составе System)'''
        sheet = book['Кнопки']  # .worksheets[10]
        cells = sheet['A1': 'C' + str(sheet.max_row)]
        sl_CPU_one = is_load_btn(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                 is_f_ind(cells[0], 'Наименование параметра'),
                                 is_f_ind(cells[0], 'CPU'))

        tmp_subgroup = ''
        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_btn_cnt(sl_CPU_one, tmp_object_BTN_CNT_sig, 'Types.BTN.BTN_PLC_View')
            tmp_subgroup += Template(tmp_group).substitute(name_group='BTN', objects=tmp_line_)

        '''Уставки(в составе System)'''
        sheet = book['Уставки']  # .worksheets[5]
        cells = sheet['A1': 'AG' + str(sheet.max_row)]
        sl_CPU_one = is_load_ai_ae_set(i, cells, is_f_ind(cells[0], 'Алгоритмическое имя'),
                                       is_f_ind(cells[0], 'Наименование параметра'),
                                       is_f_ind(cells[0], 'Единицы измерения'),
                                       is_f_ind(cells[0], 'Короткое наименование'),
                                       is_f_ind(cells[0], 'Количество знаков'),
                                       is_f_ind(cells[0], 'CPU'))

        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_ai_ae_set(sl_CPU_one, tmp_object_AIAESET, 'Types.SET.SET_PLC_View')
            tmp_subgroup += Template(tmp_group).substitute(name_group='SET', objects=tmp_line_)

        '''Наработки(CNT- прочитали при ИМ) в составе System)'''
        if len(sl_cnt) != 0:
            tmp_line_ = is_create_objects_btn_cnt(sl_cnt, tmp_object_BTN_CNT_sig, 'Types.CNT.CNT_PLC_View')
            tmp_subgroup += Template(tmp_group).substitute(name_group='CNT', objects=tmp_line_)

        '''Защиты(PZ) в составе System'''
        sheet = book['Сигналы']  # .worksheets[11]
        cells = sheet['A1': 'N' + str(sheet.max_row)]
        sl_CPU_one, num_pz = is_load_pz(i, cells, num_pz,
                                        is_f_ind(cells[0], 'Наименование параметра'),
                                        is_f_ind(cells[0], 'Тип защиты'),
                                        is_f_ind(cells[0], 'Единица измерения'),
                                        is_f_ind(cells[0], 'CPU'))

        if len(sl_CPU_one) != 0:
            tmp_line_ = is_create_objects_pz(sl_CPU_one, tmp_object_PZ, 'Types.PZ.PZ_PLC_View')
            tmp_subgroup += Template(tmp_group).substitute(name_group='PZ', objects=tmp_line_)

        '''ПС(WRN) в составе System, каждая ПС как отдельный объект, дополняем словарь, созданный при анализе DI'''
        tmp_wrn, sl_ts, sl_ppu, sl_alr, sl_modes, sl_alg = is_load_sig(i, cells,
                                                                       is_f_ind(cells[0], 'Алгоритмическое имя'),
                                                                       is_f_ind(cells[0], 'Наименование параметра'),
                                                                       is_f_ind(cells[0], 'Тип защиты'),
                                                                       is_f_ind(cells[0], 'CPU'))
        sl_wrn = {**sl_wrn, **tmp_wrn}

        if len(sl_wrn) != 0:
            tmp_line_ = is_create_objects_sig(sl_wrn, tmp_object_BTN_CNT_sig)
            tmp_subgroup += Template(tmp_group).substitute(name_group='WRN', objects=tmp_line_)

        '''ТС в составе System, сам словарь загружен ранее вместе с ПС'''
        if len(sl_ts) != 0:
            tmp_line_ = is_create_objects_sig(sl_ts, tmp_object_BTN_CNT_sig)
            tmp_subgroup += Template(tmp_group).substitute(name_group='TS', objects=tmp_line_)

        '''ППУ в составе System, сам словарь загружен ранее вместе с ПС'''
        if len(sl_ppu) != 0:
            tmp_line_ = is_create_objects_sig(sl_ppu, tmp_object_BTN_CNT_sig)
            tmp_subgroup += Template(tmp_group).substitute(name_group='PPU', objects=tmp_line_)

        '''ALR(Защиты и АС) в составе System, сам словарь загружен ранее вместе с ПС'''
        if len(sl_alr) != 0:
            tmp_line_ = is_create_objects_sig(sl_alr, tmp_object_BTN_CNT_sig)
            tmp_subgroup += Template(tmp_group).substitute(name_group='ALR', objects=tmp_line_)

        '''MODES(Режимы) в составе System, сам словарь загружен ранее вместе с ПС'''
        if len(sl_modes) != 0:
            tmp_line_ = is_create_objects_sig(sl_modes, tmp_object_BTN_CNT_sig)
            tmp_subgroup += Template(tmp_group).substitute(name_group='MODES', objects=tmp_line_)

        '''ALG в составе System, сам словарь загружен ранее вместе с ПС(позже поддержать новый конфигуратор)'''
        if len(sl_alg) != 0:
            tmp_line_ = is_create_objects_sig(sl_alg, tmp_object_BTN_CNT_sig)
            tmp_subgroup += Template(tmp_group).substitute(name_group='ALG', objects=tmp_line_)

        '''Добавляем топливный регулятор, если для данного контроллера указан ТР ПС90 в настройках'''
        if 'ТР' in sl_CPU_spec[i]:
            tmp_subgroup += tmp_tr_ps90.rstrip()

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

        # print(os.path.exists(os.path.join(os.path.dirname(__file__), 'File_out')))
        if not os.path.exists(os.path.join(os.path.dirname(__file__), 'File_out')):
            os.mkdir(os.path.join(os.path.dirname(__file__), 'File_out'))
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
                '''для каждого контроллера создадим отдельный файл для импорта только его одного'''
                tmp_plc = Template(tmp_trei).substitute(plc_name='PLC_' + i + str(num_obj_plc), plc_name_type='m903e',
                                                        ip_eth1=pref_IP[0] + sl_object_all[obj][1][index_tmp],
                                                        ip_eth2=pref_IP[1] + sl_object_all[obj][1][index_tmp],
                                                        dp_app=tmp_line_)
                with open(os.path.join(os.path.dirname(__file__), 'File_out', f'file_out_plc_{i}_{num_obj_plc}.omx-export'),
                          'w', encoding='UTF-8') as f:
                    f.write(Template(tmp_global).substitute(dp_node=tmp_plc))
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

    os.remove('file_plc.txt')
    os.remove('file_out_group.txt')
    '''os.remove('file_out_objects.txt')'''
    os.remove('file_app_out.txt')
    print(datetime.datetime.now())

    create_index()
    print(datetime.datetime.now())

except (Exception, KeyError):

    logging.basicConfig(filename='app.log', filemode='a', datefmt='%d.%m.%y %H:%M:%S',
                        format='%(levelname)s - %(message)s - %(asctime)s')
    logging.exception("Ошибка выполнения")
    print('Произошла ошибка выполнения')
