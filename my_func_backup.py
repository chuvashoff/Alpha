from string import Template
def is_create_objects_ai_ae(sl_cpu, template_text, object_type):
    file_out = open('file_out_objects.txt', 'w')
    for key, value in sl_cpu.items():
        file_out.write(Template(template_text).substitute(object_name=key, object_type=object_type,
                                                          object_aspect='Types.PLC_Aspect',
                                                          text_description=value[0], text_eunit=value[1],
                                                          text_fracdigits=value[2]))
    file_out.close()
    with open('file_out_objects.txt', 'r') as f:
        tmp_line_object = f.read().rstrip()
    return tmp_line_object

def is_create_objects_di(sl_cpu, template_text, object_type):
    file_out = open('file_out_objects.txt', 'w')
    for key, value in sl_cpu.items():
        file_out.write(Template(template_text).substitute(object_name=key, object_type=object_type,
                                                          object_aspect='Types.PLC_Aspect',
                                                          text_description=value[0]))
    file_out.close()
    with open('file_out_objects.txt', 'r') as f:
        tmp_line_object = f.read().rstrip()
    return tmp_line_object