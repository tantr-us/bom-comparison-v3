from modules.errors import RefDesError
import re, string


def transform_ref_des(ref_des_str):
    if ' ' in ref_des_str:
        raise RefDesError('Reference Designators contain whitespace.')
    else:
        ref_des_list = None
        ref_des_list = ref_des_str.split(',')
        new_ref_des = []
        for ref_des in ref_des_list:
            ref_des = ref_des.strip()
            if not _has_special_char(ref_des):
                if '-' not in ref_des:
                    new_ref_des.append(ref_des)
                else:
                    if ref_des.count('-') == 1:
                        start, end = ref_des.split('-')
                        start_letter, start_number = _split_letter_and_number(start)
                        if start_number != None:
                            if re.match(r'([a-zA-Z]+)(\d+)', end):
                                end_letter, end_number = _split_letter_and_number(end)
                                for i in range(int(start_number), int(end_number)+1):
                                    new_ref_des.append(f'{start_letter}{i}')
                            elif re.match(r'\d+', end):
                                for i in range(int(start_number), int(end)+1):
                                    new_ref_des.append(f'{start_letter}{i}')
                        else:
                            new_ref_des.append(ref_des)
                    else:
                        new_ref_des.append(ref_des)
            else:
                new_ref_des.append(ref_des)
        return new_ref_des
 
def _has_special_char(ref_des):
    spec_char = string.punctuation.replace('-', '')
    found_special_char = False
    for c in spec_char:
        if c in ref_des:
            found_special_char = True
            break
    return found_special_char
    
 
def _split_letter_and_number(location):
    regex = r'([a-zA-Z]+)(\d+)'
    match = re.match(regex, location)
    if match:
        letter = match.group(1)
        number = match.group(2)
        return letter, number
    return None, None