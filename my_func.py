from string import Template
import os
'''
Читаем и грузим в словарь, где ключ - алг имя, а в значении список - русское наименование
ед. изм-я, короткое наименование и количество знаков после запятой
'''


def is_load_ai_ae_set(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller:
            tmp['_'.join(par[1].value.split('|'))] = [par[0].value, par[28].value, par[29].value, par[30].value]
    return tmp


'''
Функция для чтения обвеса DI, возможно потребуется в список грузить больше информации, пока читаем только алг имя, 
описание, цвет при наличии(color_on) и цвет при отсутствии(color_off)
'''


def is_load_di(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller and par[14].value == 'Нет':
            tmp['_'.join(par[1].value.split('|'))] = [par[0].value, par[19].fill.start_color.index,
                                                      par[20].fill.start_color.index]
    return tmp


'''Для им держим пока рус наименование, вид има, род, тип има по отображению, флаг наработки, флаг перестановки'''


def is_load_im(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller:
            tmp[par[1].value] = [par[0].value, par[5].value, par[4].value, par[19].value[0], par[14].value,
                                 par[15].value]
    return tmp


def is_load_im_ao(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller and par[3].value == 'Да':
            tmp[par[1].value] = [par[0].value, 'ИМАО', par[25].value, par[26].value[0]]
    return tmp


def is_load_btn(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller:
            tmp['BTN_' + par[1].value[par[1].value.find('|')+1:]] = [par[0].value]
    return tmp


'''Для защит держим рус имя и ед. измерения'''


def is_load_pz(controller, cell, num_pz):
    tmp = {}
    for par in cell:
        if par[0].value is None:
            break
        if par[4].value not in 'АОссАОбсВОссВОбсАОНО':
            continue
        elif par[2].value == controller and par[4].value in 'АОссАОбсВОссВОбсАОНО':
            '''обработка спецсимволов html в русском наименовании'''
            tmp_name = par[0].value.replace('<', '&lt;')
            tmp_name = tmp_name.replace('>', '&gt;')
            tmp['A' + str(num_pz).zfill(3)] = [tmp_name, par[9].value]
            num_pz += 1
    return tmp, num_pz


'''В словаре ПС держим текст и важность 40(если что сможем здесь контролировать)'''


def is_load_sig(controller, cell):
    tmp_wrn = {}
    for par in cell:
        if par[0].value is None:
            break
        if par[2].value == controller:
            if 'ПС' in par[4].value:
                tmp_wrn[par[1].value[par[1].value.find('|')+1:]] = [par[0].value, '40']

    return tmp_wrn


'''
Создаёт набор объектов возвращает его (ранне клала в промежуточный файл, теперь этого не делает, функция осталась в bk)
'''


def is_create_objects_ai_ae_set(sl_cpu, template_text, object_type):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key, object_type=object_type,
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0], text_eunit=value[1],
                                                              short_name=value[2], text_fracdigits=value[3])

    return tmp_line_object.rstrip()


sl_color_di = {'FF969696': '0', 'FF00B050': '1', 'FFFFFF00': '2', 'FFFF0000': '3'}


def is_create_objects_di(sl_cpu, template_text, object_type):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key, object_type=object_type,
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0], color_on=sl_color_di[value[1]],
                                                              color_off=sl_color_di[value[2]])

    return tmp_line_object.rstrip()


sl_im_PLC = {'ИМ1Х0': 'IM1x0.PLC_IM1x0', 'ИМ1Х1': 'IM1x1.PLC_IM1x1', 'ИМ1Х2': 'IM1x2.IM1x2_PLC',
             'ИМ2Х2': 'IM2x2.IM2x2_PLC', 'ИМ2Х4': 'IM2x2.PLC_IM2x4', 'ИМ1Х0и': 'IM1x0.PLC_IM1x0',
             'ИМ1Х1и': 'IM1x1.PLC_IM1x1', 'ИМ1Х2и': 'IM1x2.IM1x2_PLC', 'ИМ2Х2с': 'IM2x2.IM2x2_PLC',
             'ИМАО': 'IM_AO.PLC_IM_AO'}

sl_gender = {'С': '0', 'М': '1', 'Ж': '2'}


def is_create_objects_im(sl_cpu, template_text):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key,
                                                              object_type='Types.' + sl_im_PLC[value[1]],
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0], gender=sl_gender[value[2]],
                                                              start_view=value[3])

    return tmp_line_object.rstrip()


def is_create_objects_btn_cnt(sl_cpu, template_text, object_type):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key, object_type=object_type,
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0])

    return tmp_line_object.rstrip()


def is_create_objects_pz(sl_cpu, template_text, object_type):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key, object_type=object_type,
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0],
                                                              text_eunit=value[1])

    return tmp_line_object.rstrip()


'''Считываем шаблоны для формирования событий параметров'''
sl_wrn = {}
with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_onoff_json'), 'r', encoding='UTF-8') as f:
    sl_wrn['Да (по наличию)'] = f.readline().rstrip()
    sl_wrn['Да (по отсутствию)'] = f.readline().rstrip()


def is_create_for_types_wrn(sl_cpu, template_text):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(socket_par_name=key, socket_par_type='bool',
                                                              json_message=Template(sl_wrn['Да (по наличию)']).substitute(text_description=value[0], par_severity=value[1]),
                                                              text_description=value[0])
    return tmp_line_object.rstrip()
