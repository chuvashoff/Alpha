from string import Template

'''Функция для поиска индекса нужно столбца'''


def is_f_ind(cell, name_col):
    for i in range(len(cell)):
        if cell[i].value.replace('\n', ' ') == name_col:
            return i
    return 0


'''Функция для замены в строке спецсимволов HTML'''


def is_cor_chr(st):
    sl_chr = {'<': '&lt;', '>': '&gt;', '"': '&quot;'}
    tmp = list()
    tmp.extend(st)
    for i in range(len(tmp)):
        if tmp[i] in sl_chr:
            tmp[i] = sl_chr[tmp[i]]
    return ''.join(tmp)


'''
Читаем и грузим в словарь, где ключ - алг имя, а в значении список - русское наименование
ед. изм-я, короткое наименование и количество знаков после запятой
'''


def is_load_ai_ae_set(controller, cell, alg_name, name_par, eunit, short_name, f_dig, cpu):
    tmp = {}
    for par in cell:
        if par[cpu].value == controller:
            tmp['_'.join(par[alg_name].value.split('|'))] = [is_cor_chr(par[name_par].value),
                                                             par[eunit].value, par[short_name].value, par[f_dig].value]
    return tmp


'''
Функция для чтения обвеса DI, возможно потребуется в список грузить больше информации, пока читаем только алг имя, 
описание, цвет при наличии(color_on) и цвет при отсутствии(color_off)
Также собираем словарь ПС -  текст сообщения и держим тип сообщения
'''


def is_load_di(controller, cell, alg_name, im, name_par, c_on, c_off, ps, ps_msg, cpu):
    tmp = {}
    tmp_wrn = {}
    for par in cell:
        if par[cpu].value == controller and par[im].value == 'Нет':
            tmp['_'.join(par[alg_name].value.split('|'))] = [is_cor_chr(par[name_par].value),
                                                             par[c_on].fill.start_color.index,
                                                             par[c_off].fill.start_color.index]
            if par[ps].value != 'Нет':
                tmp_wrn['_'.join(par[alg_name].value.split('|'))] = [is_cor_chr(par[ps_msg].value), par[ps].value]
    return tmp, tmp_wrn


'''Для им держим пока рус наименование, вид има, род, тип има по отображению, флаг наработки, флаг перестановки'''


def is_load_im(controller, cell, alg_name, name_par, type_im, gender, w_time, swap, cpu):
    tmp = {}
    for par in cell:
        if par[cpu].value == controller:
            tmp[par[alg_name].value] = [is_cor_chr(par[name_par].value), par[type_im].value,
                                        par[gender].value, par[19].value[0], par[w_time].value,
                                        par[swap].value]
    return tmp


def is_load_im_ao(controller, cell, alg_name, name_par, gender, im, cpu):
    tmp = {}
    for par in cell:
        if par[cpu].value == controller and par[im].value == 'Да':
            tmp[par[alg_name].value] = [is_cor_chr(par[name_par].value), 'ИМАО', par[gender].value, par[26].value[0]]
    return tmp


def is_load_btn(controller, cell, alg_name, name_par, cpu):
    tmp = {}
    for par in cell:
        if par[cpu].value == controller:
            tmp['BTN_' + par[alg_name].value[par[alg_name].value.find('|')+1:]] = [is_cor_chr(par[name_par].value)]
    return tmp


'''Для защит держим рус имя и ед. измерения'''


def is_load_pz(controller, cell, num_pz, par_name, type_protect, eunit, cpu):
    tmp = {}
    for par in cell:
        if par[par_name].value is None:
            break
        if par[type_protect].value not in 'АОссАОбсВОссВОбсАОНО':
            continue
        elif par[cpu].value == controller and par[type_protect].value in 'АОссАОбсВОссВОбсАОНО':
            '''обработка спецсимволов html в русском наименовании'''
            '''
            tmp_name = par[par_name].value.replace('<', '&lt;')
            tmp_name = tmp_name.replace('>', '&gt;')
            '''
            if par[eunit].value == '-999.0':
                tmp_eunit = str(par[eunit].comment)[str(par[eunit].comment).find(' ')+1:
                                                    str(par[eunit].comment).find('by')]
            else:
                tmp_eunit = par[eunit].value
            tmp['A' + str(num_pz).zfill(3)] = [par[type_protect].value + '. ' + is_cor_chr(par[par_name].value),
                                               tmp_eunit]
            num_pz += 1
    return tmp, num_pz


'''словарь ПС -  текст сообщения и держим тип сообщения'''


def is_load_sig(controller, cell, alg_name, par_name, type_protect, cpu):
    tmp_wrn = {}
    tmp_alr = {}
    tmp_ts = {}
    tmp_ppu = {}
    tmp_modes = {}
    for par in cell:
        if par[par_name].value is None:
            break
        if par[cpu].value == controller:
            if 'ПС' in par[type_protect].value:
                tmp_wrn[par[alg_name].value[par[alg_name].value.find('|')+1:]] = [is_cor_chr(par[par_name].value),
                                                                                  'Да (по наличию)']
            elif par[type_protect].value in 'АОссАОбсВОссВОбсАОНО' or 'АС' in par[type_protect].value:
                if 'АС' in par[type_protect].value:
                    tmp_alr[par[alg_name].value[par[alg_name].value.find('|') + 1:]] = ['АС. ' +
                                                                                        is_cor_chr(par[par_name].value),
                                                                                        'АС']
                else:
                    tmp_alr[par[alg_name].value[par[alg_name].value.find('|') + 1:]] = [par[type_protect].value + '. ' +
                                                                                        is_cor_chr(par[par_name].value),
                                                                                        'Защита']
            elif 'ТС' in par[type_protect].value:
                tmp_ts[par[alg_name].value[par[alg_name].value.find('|')+1:]] = [is_cor_chr(par[par_name].value), 'ТС']
            elif par[type_protect].value in ('ГР', 'ХР'):
                tmp_ppu[par[alg_name].value[par[alg_name].value.find('|') + 1:]] = [is_cor_chr(par[par_name].value),
                                                                                    'ППУ']
            elif 'Режим' in par[type_protect].value:
                tmp_modes[par[alg_name].value[par[alg_name].value.find('|') + 1:]] = ['Режим &quot;' +
                                                                                      is_cor_chr(par[par_name].value) +
                                                                                      '&quot;', 'Режим']
                tmp_modes['regNum'] = ['Номер режима', 'Номер режима']
    return tmp_wrn, tmp_ts, tmp_ppu, tmp_alr, tmp_modes


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


sl_im_PLC = {'ИМ1Х0': 'IM1x0.IM1x0_PLC_View', 'ИМ1Х1': 'IM1x1.IM1x1_PLC_View', 'ИМ1Х2': 'IM1x2.IM1x2_PLC_View',
             'ИМ2Х2': 'IM2x2.IM2x2_PLC_View', 'ИМ2Х4': 'IM2x2.IM2x4_PLC_View', 'ИМ1Х0и': 'IM1x0.IM1x0_PLC_View',
             'ИМ1Х1и': 'IM1x1.IM1x1_PLC_View', 'ИМ1Х2и': 'IM1x2.IM1x2_PLC_View', 'ИМ2Х2с': 'IM2x2.IM2x2_PLC_View',
             'ИМАО': 'IM_AO.IM_AO_PLC_View'}

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


sl_type_sig = {'Да (по наличию)': 'Types.WRN_On.WRN_On_PLC_View',
               'Да (по отсутствию)': 'Types.WRN_Off.WRN_Off_PLC_View',
               'ТС': 'Types.TS.TS_PLC_View',
               'ППУ': 'Types.PPU.PPU_PLC_View',
               'Защита': 'Types.ALR.ALR_PLC_View',
               'АС': 'Types.ALR.ALR_PLC_View',
               'Режим': 'Types.MODES.MODES_PLC_View',
               'Номер режима': 'Types.MODES.regNum_PLC_View'}


def is_create_objects_sig(sl_cpu, template_text):
    tmp_line_object = ''
    for key, value in sl_cpu.items():
        tmp_line_object += Template(template_text).substitute(object_name=key,
                                                              object_type=sl_type_sig[value[1]],
                                                              object_aspect='Types.PLC_Aspect',
                                                              text_description=value[0])
    return tmp_line_object.rstrip()
