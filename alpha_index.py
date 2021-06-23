import os
import re
from string import Template


def create_sl(text, str_check):
    sl_tmp = {}
    for i in text:
        if str_check in i and '//' not in i and str_check.replace('_', '|') not in i:
            if 'FAST|' in i:
                sl_tmp[i[:i.find(' ')]] = int(i[i.rfind('[') + 1:i.rfind(']')])
            else:
                sl_tmp[i[:i.find('(')]] = int(i[i.rfind('(') + 1:i.rfind(')')])

    sl_tmp = {key: value for key, value in sl_tmp.items() if f'FAST|{key}' not in sl_tmp}
    '''В словаре sl_tmp лежит индекс массива: алг имя (в том числе FAST|+)'''
    sl_tmp = {value: key for key, value in sl_tmp.items()}
    return sl_tmp


def create_sl_im(text):
    sl_tmp = {}
    cnt_set = set()
    for i in text:
        if 'IM|' in i and '//' not in i and '[' in i:
            a = i.split(':=')[0].strip()
            sl_tmp[int(i[i.rfind('[') + 1:i.rfind(']')])] = a[a.find('|')+1:a.rfind('_')]
        if 'IM|' in i and '//' not in i and ('WorkTime' in i or 'Swap' in i) and 'TCycle' not in i:
            aa = i.split(':=')[0].strip()
            cnt_set.add(aa[aa.find('|')+1:])
    '''В словаре sl_tmp лежит индекс массива: алг имя; в cnt_set лежат используемые наработки'''
    return sl_tmp, cnt_set


def create_sl_pz(text):
    tmp_set = set()
    for i in text:
        if 'FAST|ALR_' in i and '//' not in i:
            a = i.split(':=')[0].strip()
            tmp_set.add(a[a.find('_') + 1:])
    '''в множестве tmp_set лежат используемые аварии'''
    return tmp_set


def create_group_par(sl_global_par, sl_local_par, sl_global_fast, template_arc_index, template_no_arc_index,
                     pref_par, source):
    sl_data_cat = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Discrete'
    }
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    for key, value in sl_global_par.items():
        tmp_i = int(key[key.find('[')+1:key.find(']')])
        tmp_sub_name = key[:key.find('[')]
        if tmp_i not in sl_local_par:
            continue
        if 'FAST|' in sl_local_par[tmp_i] and re.fullmatch(r'Value', key[:key.find('[')]):
            value[0] = sl_global_fast[sl_local_par[tmp_i]]
            a = sl_local_par[tmp_i][sl_local_par[tmp_i].find('|')+1:]
            temp = template_arc_index
            pref_arc = 'Arc'
        else:
            a = sl_local_par[tmp_i][sl_local_par[tmp_i].find('|')+1:]
            temp = template_no_arc_index
            pref_arc = f'NoArc{sl_data_cat[value[1]]}'
        s_out += Template(temp).substitute(name_signal=f'{pref_par}.{a}.{tmp_sub_name}', type_signal=sl_type[value[1]],
                                           index=value[0], data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_group_im(sl_global_im, sl_local_im, sl_global_fast, template_arc_index, template_no_arc_index, source):
    sl_data_cat = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Discrete'
    }
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    for key, value in sl_global_im.items():
        tmp_i = int(key[key.find('[') + 1:key.find(']')])
        tmp_sub_name = key[:key.find('[')]
        if tmp_i not in sl_local_im:
            continue
        if f'FAST|IM_{sl_local_im[tmp_i]}_{tmp_sub_name}' in sl_global_fast:
            value[0] = sl_global_fast[f'FAST|IM_{sl_local_im[tmp_i]}_{tmp_sub_name}']
            a = sl_local_im[tmp_i]
            temp = template_arc_index
            pref_arc = 'Arc'
        elif f'FAST|AO_{sl_local_im[tmp_i]}_{tmp_sub_name}' in sl_global_fast:
            value[0] = sl_global_fast[f'FAST|AO_{sl_local_im[tmp_i]}_{tmp_sub_name}']
            a = sl_local_im[tmp_i]
            temp = template_arc_index
            pref_arc = 'Arc'
        else:
            a = sl_local_im[tmp_i]
            temp = template_no_arc_index
            pref_arc = f'NoArc{sl_data_cat[value[1]]}'
        s_out += Template(temp).substitute(name_signal=f'IM.{a}.{tmp_sub_name}', type_signal=sl_type[value[1]],
                                           index=value[0], data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_group_btn(sl_global_btn, sl_local_btn, template_no_arc_index, source):
    sl_data_cat = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Discrete'
    }
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    for key, value in sl_global_btn.items():
        tmp_i = int(key[key.find('[') + 1:key.find(']')])
        tmp_sub_name = key[:key.find('[')]
        if tmp_i not in sl_local_btn:
            continue
        a = sl_local_btn[tmp_i]
        pref_arc = f'NoArc{sl_data_cat[value[1]]}'
        s_out += Template(template_no_arc_index).substitute(name_signal=f'System.BTN.{a}.{tmp_sub_name}',
                                                            type_signal=sl_type[value[1]], index=value[0],
                                                            data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_group_system_sig(sub_name, sl_global_sig, template_no_arc_index, source):
    sl_data_cat = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Discrete'
    }
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    for key, value in sl_global_sig.items():
        a = key
        pref_arc = f'NoArc{sl_data_cat[value[1]]}'  # для проекта Бованенково убрать '.Value'
        s_out += Template(template_no_arc_index).substitute(name_signal=f'System.{sub_name}.{a}.Value',
                                                            type_signal=sl_type[value[1]], index=value[0],
                                                            data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_group_tr(sl_global, template_no_arc_index, pref_par, source):
    sl_data_cat = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Discrete'
    }
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    for key, value in sl_global.items():
        a = key
        pref_arc = f'NoArc{sl_data_cat[value[1]]}'
        s_out += Template(template_no_arc_index).substitute(name_signal=f'{pref_par}.{a}',
                                                            type_signal=sl_type[value[1]], index=value[0],
                                                            data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_group_alr(sl_global_alr, template_arc_index, source):
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    for key, value in sl_global_alr.items():
        a = key
        pref_arc = 'Arc'  # для проекта Бованенково убрать '.Value'
        s_out += Template(template_arc_index).substitute(name_signal=f'System.ALR.{a}.Value',
                                                         type_signal=sl_type[value[1]], index=value[0],
                                                         data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_group_pz(sl_global_pz, lst_pz, sl_pz_anum, template_no_arc_index, source):
    sl_data_cat = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Discrete'
    }
    sl_type = {
        'R': 'Analog',
        'I': 'Analog',
        'B': 'Bool'
    }
    s_out = ''
    sl_pz_i = {}
    for i in range(len(sl_global_pz)//len(lst_pz)):
        lst_tmp = []  # В lst_tmp кладём списки [инд. переменной, тип ] в сответствии с lst_pz -нужные подимена
        for j in lst_pz:
            lst_tmp.append(sl_global_pz[f'{j}[{i}]'])
        sl_pz_i[i] = dict(zip(lst_pz, lst_tmp))  # {индекс массива: {подимя: [инд. переменной, тип ]}}
    lst_anum = ['A' + str(i).zfill(3) for i in range(sl_pz_anum[0], sl_pz_anum[1]+1)]
    sl_pz = dict(zip(lst_anum, sl_pz_i.values()))  # {Имя переменной(A000): {подимя: [инд. переменной, тип ]}}
    for key, value in sl_pz.items():
        for key_subname, value_subname in value.items():
            pref_arc = f'NoArc{sl_data_cat[value_subname[1]]}'
            s_out += Template(template_no_arc_index).substitute(name_signal=f'System.PZ.{key}.{key_subname}',
                                                                type_signal=sl_type[value_subname[1]],
                                                                index=value_subname[0],
                                                                data_category=f'DataCategory_{source}_{pref_arc}')
    return s_out


def create_index(lst_alg, lst_mod, lst_ppu, lst_ts, lst_wrn, sl_pz_anum, sl_cpu_spec):
    # Считываем шаблоны для карты
    with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_map_index_Arc'), 'r', encoding='UTF-8') as f:
        tmp_ind_arc = f.read()
    with open(os.path.join(os.path.dirname(__file__), 'Template', 'Temp_map_index_noArc'), 'r', encoding='UTF-8') as f:
        tmp_ind_no_arc = f.read()

    lst_ae = (
        'coSim', 'coRepair', 'coUpdRepair', 'Message.msg_fBreak', 'Message.msg_qbiValue', 'Value', 'fValue', 'iValue',
        'qbiValue', 'sChlType', 'sEnRepair', 'TRepair', 'Value100', 'aH', 'aL', 'fParam', 'fOverlow', 'fOverhigh',
        'fHighspeed', 'Repair', 'Sim', 'wH', 'wL', 'sEnaH', 'sBndaH', 'sEnaL', 'sBndaL', 'sEnwH', 'sBndwH', 'sEnwL',
        'sBndwL', 'sEnRate', 'sBndRate', 'sHighLim', 'sHys', 'sLowLim', 'sSimValue', 'sTaH', 'sTaL', 'sTRepair', 'sTwH',
        'sTwL'
    )
    lst_ai = lst_ae + ('sHighiValue', 'sLowiValue')
    lst_di = (
        'coSim', 'brk', 'kz', 'coRepair', 'coUpdRepair', 'fParam', 'Value', 'fValue', 'Message.msg_brk',
        'Message.msg_qbiValue', 'Message.msg_kz', 'qbiValue', 'Repair', 'sBlkPar', 'Sim', 'sSimValue', 'TRepair',
        'sTRepair'
    )
    lst_im1x0 = (
        'coOn', 'coOff', 'Message.msg_fwcOn', 'Message.msg_fwsDu', 'Message.msg_qbiDu', 'Message.msg_qboOn', 'oOff',
        'oOn', 'fFlt', 'fwcOn', 'pcoMan', 'pcoOff', 'pcoOn', 'pMan', 'qbiDu', 'qboOn', 'Repair', 'fwsDu', 'wbcaOff',
        'wbcaOn', 'wMU', 'wRU', 'TRepair', 'coRepair', 'sTRepair', 'coUpdRepair', 'coMan'
    )
    lst_im1x1 = (
        'coOn', 'coOff', 'coRst', 'Message.msg_fwcOn', 'Message.msg_qbiOn', 'Message.msg_fwsOn', 'Message.msg_fwsDu',
        'Message.msg_qbiDu', 'Message.msg_qboOn', 'fwcOn', 'pcoMan', 'pcoOff', 'pcoOn', 'pcoRst', 'pMan', 'qbiDu',
        'qbiOn', 'qboOn', 'stOn', 'stOff', 'stOng', 'stOffg', 'fwsDu', 'fwsOn', 'wNotOn', 'wNotOff', 'TRepair',
        'oOn', 'oOff', 'fFlt', 'Repair', 'wbcaOff', 'wbcaOn', 'wMU', 'wRU', 'sTOn_Off', 'sTRepair', 'coRepair',
        'coUpdRepair', 'coMan'
    )
    lst_im1x2 = lst_im1x1 + ('wUnkw', 'wDbl', 'Message.msg_fwsOff', 'Message.msg_qbiOff', 'fwsOff', 'qbiOff')
    lst_im2x2 = lst_im1x2 + ('coStop', 'Message.msg_fwcOff', 'Message.msg_qboOff', 'fwcOff', 'pcoStop', 'qboOff',
                             'sTcom', 'sTPress')
    lst_im_ao = (
        'Message.msgqbiPos', 'Message.msgqboSet', 'fFlt', 'pcoMan', 'pMan', 'qbiPos', 'Repair', 'wRU', 'TRepair',
        'stWay', 'fPos', 'fSet', 'ifOut', 'ifPos', 'iPos', 'Out', 'qboSet', 'Set', 'stOn', 'stOff', 'coMan', 'coRepair',
        'sTRepair', 'coUpdRepair', 'coSet', 'sHighiValue', 'sLowiValue', 'sHighLim', 'sLowLim'
    )
    lst_btn = (
        'ceOn', 'coOn', 'caLhBTN', 'pBTN', 'pcoMan'
    )
    lst_pz = (
        'ALR', 'CHECK', 'PZH', 'PZM', 'PZS', 'TDELAY', 'TRDELAY', 'WASFIRST', 'CHECKVALUE', 'SETPOINT', 'VALUE',
        'coSbros', 'BLOCKED'
    )
    lst_set = (
        'ceSP', 'seSP', 'soSP', 'Value', 'caRegOn', 'fParam', 'NoLink', 'sHighLim', 'sLowLim', 'vPar'
    )

    with open('Source_list_plc', 'r', encoding='UTF-8') as f:
        while 8:
            line_source = f.readline().strip().split(',')
            if line_source == ['']:
                break
            sl_global_ai, sl_tmp_ai = {}, {}
            sl_global_ae, sl_tmp_ae = {}, {}
            sl_global_di, sl_tmp_di, sl_wrn_di = {}, {}, {}  # sl_wrn_di - для сбора ПС по Дискретам
            sl_global_im1x0, sl_tmp_im1x0 = {}, {}
            sl_global_im1x1, sl_tmp_im1x1 = {}, {}
            sl_global_im1x2, sl_tmp_im1x2 = {}, {}
            sl_global_im2x2, sl_tmp_im2x2 = {}, {}
            sl_global_im_ao, sl_tmp_im_ao = {}, {}
            sl_global_btn, sl_tmp_btn = {}, {}
            sl_global_set, sl_tmp_set = {}, {}
            sl_global_cnt = {}
            set_all_im = set()  # sl_all_im - для ИМ с наработками
            set_cnt_im1x0, set_cnt_im1x1, set_cnt_im1x2, set_cnt_im2x2 = set(), set(), set(), set()
            sl_global_alr = {}
            set_tmp_alr = set()
            sl_global_alg = {}
            sl_global_mod = {}
            sl_global_ppu = {}
            sl_global_ts = {}
            sl_global_wrn = {}
            sl_global_pz = {}
            sl_global_fast = {}
            s_all = ''

            lst_tr_par = []
            sl_global_tr = {}

            if 'ТР' in sl_cpu_spec[line_source[0]]:
                with open(os.path.join(os.path.dirname(__file__), 'Template', 'TR_par'), 'r', encoding='UTF-8') as f_tr:
                    lst_tr_par = f_tr.read().split('\n')

            '''Если есть файл аналогов'''
            if os.path.isfile(os.path.join(line_source[1], '0_par_A.st')):
                with open(os.path.join(line_source[1], '0_par_A.st')) as f_par_a:
                    text = f_par_a.read().split('\n')
                sl_tmp_ai = create_sl(text, 'AI_')

            '''Если есть файл расчётных'''
            if os.path.isfile(os.path.join(line_source[1], '0_par_Evl.st')):
                with open(os.path.join(line_source[1], '0_par_Evl.st')) as f_par_evl:
                    text = f_par_evl.read().split('\n')
                sl_tmp_ae = create_sl(text, 'AE_')

            '''Если есть файл дискретных'''
            if os.path.isfile(os.path.join(line_source[1], '0_par_D.st')):
                with open(os.path.join(line_source[1], '0_par_D.st')) as f_par_d:
                    text = f_par_d.read().split('\n')
                sl_tmp_di = create_sl(text, 'DI_')

            '''Если есть файл ИМ_1x0'''
            if os.path.isfile(os.path.join(line_source[1], '0_IM_1x0.st')):
                with open(os.path.join(line_source[1], '0_IM_1x0.st')) as f_im:
                    text = f_im.read().split('\n')
                sl_tmp_im1x0, set_cnt_im1x0 = create_sl_im(text)

            '''Если есть файл ИМ_1x1'''
            if os.path.isfile(os.path.join(line_source[1], '0_IM_1x1.st')):
                with open(os.path.join(line_source[1], '0_IM_1x1.st')) as f_im:
                    text = f_im.read().split('\n')
                sl_tmp_im1x1, set_cnt_im1x1 = create_sl_im(text)

            '''Если есть файл ИМ_1x2'''
            if os.path.isfile(os.path.join(line_source[1], '0_IM_1x2.st')):
                with open(os.path.join(line_source[1], '0_IM_1x2.st')) as f_im:
                    text = f_im.read().split('\n')
                sl_tmp_im1x2, set_cnt_im1x2 = create_sl_im(text)

            '''Если есть файл ИМ_2x2'''
            if os.path.isfile(os.path.join(line_source[1], '0_IM_2x2.st')):
                with open(os.path.join(line_source[1], '0_IM_2x2.st')) as f_im:
                    text = f_im.read().split('\n')
                sl_tmp_im2x2, set_cnt_im2x2 = create_sl_im(text)

            '''Если есть файл ИМ_АО'''
            if os.path.isfile(os.path.join(line_source[1], '0_IM_AO.st')):
                with open(os.path.join(line_source[1], '0_IM_AO.st')) as f_im:
                    text = f_im.read().split('\n')
                sl_tmp_im_ao, bla_ = create_sl_im(text)

            '''Если есть файл кнопок'''
            if os.path.isfile(os.path.join(line_source[1], '0_BTN.st')):
                with open(os.path.join(line_source[1], '0_BTN.st')) as f_btn:
                    text = f_btn.read().split('\n')
                for i in text:
                    if 'BTN_' in i and '(' in i:
                        a = i.split(',')[0]
                        sl_tmp_btn[int(a[a.find('(')+1:])] = a[:a.find('(')]

            '''Если есть файл защит PZ'''
            if os.path.isfile(os.path.join(line_source[1], '0_PZ.st')):
                with open(os.path.join(line_source[1], '0_PZ.st')) as f_pz:
                    text = f_pz.read().split('\n')
                set_tmp_alr = create_sl_pz(text)

            '''Если есть файл уставок'''
            if os.path.isfile(os.path.join(line_source[1], '0_Par_Set.st')):
                with open(os.path.join(line_source[1], '0_Par_Set.st')) as f_set:
                    text = f_set.read().split('\n')
                sl_tmp_set = create_sl(text, 'SP_')

            '''Если есть глобальный словарь'''
            if os.path.isfile(os.path.join(line_source[1], 'global0.var')):
                with open(os.path.join(line_source[1], 'global0.var')) as f_global:
                    while 8:
                        line = f_global.readline().strip()
                        if not line:
                            break
                        if 'A_INP|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_ai[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                               line[1]]
                            else:
                                sl_global_ai['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                  int(line[10])),
                                                                                              line[1]]

                        elif 'A_EVL|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_ae[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                               line[1]]
                            else:
                                sl_global_ae['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                  int(line[10])),
                                                                                              line[1]]
                        elif 'D_INP|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_di[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                               line[1]]
                            else:
                                sl_global_di['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                  int(line[10])),
                                                                                              line[1]]
                        elif 'IM_1x0|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_im1x0[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                                  line[1]]
                            else:
                                sl_global_im1x0['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                     int(line[10])),
                                                                                                 line[1]]
                        elif 'IM_1x1|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_im1x1[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                                  line[1]]
                            else:
                                sl_global_im1x1['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                     int(line[10])),
                                                                                                 line[1]]
                        elif 'IM_1x2|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_im1x2[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                                  line[1]]
                            else:
                                sl_global_im1x2['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                     int(line[10])),
                                                                                                 line[1]]
                        elif 'IM_2x2|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_im2x2[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                                  line[1]]
                            else:
                                sl_global_im2x2['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                     int(line[10])),
                                                                                                 line[1]]
                        elif 'IM_AO|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if 'msg' not in line[0]:
                                sl_global_im_ao[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                                  line[1]]
                            else:
                                sl_global_im_ao['Message.' + line[0][line[0].find('|') + 1:]] = [max(int(line[9]),
                                                                                                     int(line[10])),
                                                                                                 line[1]]
                        elif 'BTN|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_btn[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                            line[1]]

                        elif 'IM|' in line and len(line.split(',')) >= 10 and 'WorkTime' in line or 'Swap' in line:
                            line = line.split(',')
                            sl_global_cnt[line[0][line[0].find('|')+1:]] = [max(int(line[9]), int(line[10])),
                                                                            line[1]]
                        elif 'FAST|ALR_' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_alr[line[0][line[0].find('_')+1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'ALG|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_alg[f"ALG_{line[0][line[0].find('|')+1:]}"] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'GRH|' in line and len(line.split(',')) >= 10:  # в новом конфигураторе - в ветку GRH.
                            line = line.split(',')
                            sl_global_alg[f"GRH_{line[0][line[0].find('|')+1:]}"] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'MOD|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_mod[line[0][line[0].find('|') + 1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'PPU|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_ppu[line[0][line[0].find('|') + 1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'TS|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_ts[line[0][line[0].find('|') + 1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'WRN|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_wrn[line[0][line[0].find('|') + 1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'MOD_PZ|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_pz[line[0][line[0].find('|') + 1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif 'A_SET|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            sl_global_set[line[0][line[0].find('|') + 1:]] = [max(int(line[9]), int(line[10])), line[1]]

                        elif ('svk2|' in line or 'svk|' in line) and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if f"SVK.{line[0][line[0].find('|') + 1:]}" in lst_tr_par:
                                sl_global_tr[f"SVK.{line[0][line[0].find('|') + 1:]}"] = [max(int(line[9]),
                                                                                              int(line[10])), line[1]]

                        elif 'dis|' in line and len(line.split(',')) >= 10:
                            line = line.split(',')
                            if f"DIS.{line[0][line[0].find('|') + 1:]}" in lst_tr_par:
                                sl_global_tr[f"DIS.{line[0][line[0].find('|') + 1:]}"] = [max(int(line[9]),
                                                                                              int(line[10])), line[1]]

                        if 'FAST|' in line:
                            line = line.split(',')
                            '''В словаре sl_global_fast лежит  алг имя(FAST|): индекс переменной'''
                            sl_global_fast[line[0][1:]] = max(int(line[9]), int(line[10]))
            '''В словаре sl_global_ai лежит подимя[индекс массива]: [индекс переменной, тип переменной(I, B, R)]'''
            sl_global_ai = {key: value for key, value in sl_global_ai.items() if key[:key.find('[')] in lst_ai}
            sl_global_ae = {key: value for key, value in sl_global_ae.items() if key[:key.find('[')] in lst_ae}
            for key, value in sl_global_di.items():
                if 'ValueAlg' == key[:key.find('[')]:
                    var_ = sl_tmp_di.get(int(key[key.find('[')+1:key.find(']')]), 'bla')
                    if f"DI_{var_[var_.find('_')+1:]}" in lst_wrn:
                        sl_wrn_di[f"DI_{var_[var_.find('_')+1:]}"] = value

            sl_global_di = {key: value for key, value in sl_global_di.items() if key[:key.find('[')] in lst_di}
            sl_global_im1x0 = {key: value for key, value in sl_global_im1x0.items() if key[:key.find('[')] in lst_im1x0}
            sl_global_im1x1 = {key: value for key, value in sl_global_im1x1.items() if key[:key.find('[')] in lst_im1x1}
            sl_global_im1x2 = {key: value for key, value in sl_global_im1x2.items() if key[:key.find('[')] in lst_im1x2}
            sl_global_im2x2 = {key: value for key, value in sl_global_im2x2.items() if key[:key.find('[')] in lst_im2x2}
            sl_global_im_ao = {key: value for key, value in sl_global_im_ao.items() if key[:key.find('[')] in lst_im_ao}
            sl_global_btn = {key: value for key, value in sl_global_btn.items() if key[:key.find('[')] in lst_btn}
            sl_global_pz = {key: value for key, value in sl_global_pz.items() if key[:key.find('[')] in lst_pz}
            sl_global_set = {key: value for key, value in sl_global_set.items() if key[:key.find('[')] in lst_set}
            # отсуда и далее до нар ориентируемся на то, что есть в конфигураторе, так как в проекте может быть "мусор"
            # в данных словарях лежит alg имя: [индекс переменной, тип переменной(I, B, R)]
            sl_global_alg = {key: value for key, value in sl_global_alg.items() if key in lst_alg}
            sl_global_mod = {key: value for key, value in sl_global_mod.items() if key in lst_mod}
            sl_global_ppu = {key: value for key, value in sl_global_ppu.items() if key in lst_ppu}
            sl_global_ts = {key: value for key, value in sl_global_ts.items() if key in lst_ts}
            sl_global_wrn = {key: value for key, value in sl_global_wrn.items() if key in lst_wrn}
            sl_global_wrn.update(sl_wrn_di)

            # print(line_source[0], len(sl_global_pz))
            '''Объединяем множества ИМ в одно для накопления наработок'''
            for jj in [set_cnt_im1x0, set_cnt_im1x1, set_cnt_im1x2, set_cnt_im2x2]:
                set_all_im.update(jj)
            sl_global_cnt = {key: value for key, value in sl_global_cnt.items() if key in set_all_im}

            sl_global_alr = {key: value for key, value in sl_global_alr.items() if key in set_tmp_alr}

            '''Обработка и запись в карту аналогов'''

            if sl_global_ai and sl_tmp_ai:
                s_all += create_group_par(sl_global_ai, sl_tmp_ai, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                          'AI', line_source[0])

            '''Обработка и запись в карту расчётных'''

            if sl_global_ae and sl_tmp_ae:
                s_all += create_group_par(sl_global_ae, sl_tmp_ae, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                          'AE', line_source[0])

            '''Обработка и запись в карту дискретных'''

            if sl_global_di and sl_tmp_di:
                s_all += create_group_par(sl_global_di, sl_tmp_di, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                          'DI', line_source[0])

            '''Обработка и запись в карту ИМ1x0'''

            if sl_global_im1x0 and sl_tmp_im1x0:
                s_all += create_group_im(sl_global_im1x0, sl_tmp_im1x0, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                         line_source[0])

            '''Обработка и запись в карту ИМ1x1'''

            if sl_global_im1x1 and sl_tmp_im1x1:
                s_all += create_group_im(sl_global_im1x1, sl_tmp_im1x1, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                         line_source[0])

            '''Обработка и запись в карту ИМ1x2'''

            if sl_global_im1x2 and sl_tmp_im1x2:
                s_all += create_group_im(sl_global_im1x2, sl_tmp_im1x2, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                         line_source[0])

            '''Обработка и запись в карту ИМ2x2'''

            if sl_global_im2x2 and sl_tmp_im2x2:
                s_all += create_group_im(sl_global_im2x2, sl_tmp_im2x2, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                         line_source[0])

            '''Обработка и запись в карту ИМ_АО'''

            if sl_global_im_ao and sl_tmp_im_ao:
                s_all += create_group_im(sl_global_im_ao, sl_tmp_im_ao, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                         line_source[0])

            '''Обработка и запись в карту кнопок'''

            if sl_global_btn and sl_tmp_btn:
                s_all += create_group_btn(sl_global_btn, sl_tmp_btn, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту наработок'''

            if sl_global_cnt:
                s_all += create_group_system_sig('CNT', sl_global_cnt, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту ALR'''

            if sl_global_alr:
                s_all += create_group_alr(sl_global_alr, tmp_ind_arc, line_source[0])

            '''Обработка и запись в карту ALG'''

            if sl_global_alg:
                s_all += create_group_system_sig('ALG', sl_global_alg, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту Режимы'''

            if sl_global_mod:
                s_all += create_group_system_sig('MODES', sl_global_mod, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту PPU'''

            if sl_global_ppu:
                s_all += create_group_system_sig('PPU', sl_global_ppu, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту TS'''

            if sl_global_ts:
                s_all += create_group_system_sig('TS', sl_global_ts, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту WRN'''

            if sl_global_wrn:
                s_all += create_group_system_sig('WRN', sl_global_wrn, tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту Защит(PZ)'''

            if sl_global_pz and len(sl_pz_anum[line_source[0]]) != 1:
                s_all += create_group_pz(sl_global_pz, lst_pz, sl_pz_anum[line_source[0]], tmp_ind_no_arc, line_source[0])

            '''Обработка и запись в карту уставок'''
            if sl_global_set and sl_tmp_set:
                s_all += create_group_par(sl_global_set, sl_tmp_set, sl_global_fast, tmp_ind_arc, tmp_ind_no_arc,
                                          'System.SET', line_source[0])

            '''Обработка и запись в карту ТР'''
            if sl_global_tr:
                s_all += create_group_tr(sl_global_tr, tmp_ind_no_arc, 'System.TR', line_source[0])

            with open(f'trei_map_{line_source[0]}.xml', 'w') as f_out:
                f_out.write('<root format-version=\"0\">\n' + s_all.rstrip() + '\n</root>')
