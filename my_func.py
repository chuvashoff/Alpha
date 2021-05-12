from string import Template
'''
Читаем и грузим в словарь, где ключ - алг имя, а в значении список - русское наименование
ед. изм-я и количество знаков после запятой
'''


def is_load_ai_ae(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller:
            tmp['_'.join(par[1].value.split('|'))] = [par[0].value, par[28].value, par[30].value]
    return tmp


'''
Функция для чтения обвеса DI, возможно потребуется в список грузить больше информации, пока читаем только алг имя и 
описание
'''


def is_load_di(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller and par[14].value == 'Нет':
            tmp['_'.join(par[1].value.split('|'))] = [par[0].value]
    return tmp


def is_load_im(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller:
            tmp[par[1].value] = [par[0].value, par[5].value]
    return tmp


def is_load_im_ao(controller, cell):
    tmp = {}
    for par in cell:
        if par[2].value == controller and par[3].value == 'Да':
            tmp[par[1].value] = [par[0].value, 'ИМАО']
    return tmp


'''
Создаёт набор объектов возвращает его (ранне клала в промежуточный файл, теперь этого не делает, функция осталась в bk)
'''


def is_create_objects_ai_ae(sl_cpu, template_text, object_type):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key, object_type=object_type,
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0], text_eunit=value[1],
                                                              text_fracdigits=value[2])

    return tmp_line_object.rstrip()


def is_create_objects_di(sl_cpu, template_text, object_type):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key, object_type=object_type,
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0])

    return tmp_line_object.rstrip()


sl_im_PLC = {'ИМ1Х0': 'IM1x0.PLC_IM1x0', 'ИМ1Х1': 'IM1x1.PLC_IM1x1', 'ИМ1Х2': 'IM1x2.IM1x2_PLC',
             'ИМ2Х2': 'IM2x2.IM2x2_PLC', 'ИМ2Х4': 'IM2x2.PLC_IM2x4', 'ИМ1Х0и': 'IM1x0.PLC_IM1x0',
             'ИМ1Х1и': 'IM1x1.PLC_IM1x1', 'ИМ1Х2и': 'IM1x2.IM1x2_PLC', 'ИМ2Х2с': 'IM2x2.IM2x2_PLC',
             'ИМАО': 'IM_AO.PLC_IM_AO'}


def is_create_objects_im(sl_cpu, template_text):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key,
                                                              object_type='Types.' + sl_im_PLC[sl_cpu[key][1]],
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0])

    return tmp_line_object.rstrip()
