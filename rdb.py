import os
import shutil
import xlwings as xw


def update_template_400(word_bits):
    """Updates an SEL rdb template folder structure comprised of .txt files

    Args:
        word_bits (): list comprised of dictionaries {element:val, value:val, qs_group:val}
        """

    clear_groups = ['D1', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9',
                    'A10']
    file_names = os.listdir('.')
    # read file
    for file in file_names:
        found = []
        # read all .txt files
        if file.lower().endswith('.txt'):
            # isolate string segment after '_' and before '.TXT'
            settings_group = file.split("_")[1].split('.')[0]  # 'SET_1.TXT' -> 1 = quickset settings group
            file_handle = open(file, 'r')
            content = file_handle.readlines()
            file_handle.close()
            # traverse all excel variables and replace .txt line if there is a match
            for element in word_bits:
                for line in content:
                    if line.startswith(element['element'] + ',') and (settings_group == element['qs_group'] or
                                                                      element['qs_group'] is None):
                        index = content.index(line)
                        if element['value']:
                            content[index] = element['element'] + ',"' + str(element['value']) + '"' + '\x1c\n'
                        found.append(index)
                        break

            if settings_group in clear_groups:
                for line in content:
                    index = content.index(line)
                    if (index not in found and (len(line.split(',')) > 1)) or not element['value']:
                        temp = content[index].split(',')[0]
                        content[index] = temp + ',""\x1c\n'

            if settings_group == 'F1':
                for line in content:
                    index = content.index(line)
                    if index not in found and (line.startswith('DP_NAM') or line.startswith('DP_SIZE')):
                        temp = content[index].split(',')[0]
                        content[index] = temp + ',""\x1c\n'

            # write new content to file
            file_handle = open(file, 'w', encoding="ascii")
            for line in content:
                file_handle.write(line)
            file_handle.close()


def get_wordbits(relay, settings, pmu=True, mtr=False, dpac=False):
    """Extracts word bits from settings table

    Args:
        relay (list): relay class definition including RID, IP, Settings Class, Logic Class, etc.
        settings (list): list including word bits and their associated values and properties
        pmu (bool): include PMU station name
        """

    float_index = settings[0].index('Float')
    word_bits = []
    if mtr:
        word_bits.append({'element': 'MID', 'value': relay[0], 'qs_group': None})
    elif dpac:
        word_bits.append({'element': 'DID', 'value': relay[0], 'qs_group': None})
    else:
        word_bits.append({'element': 'RID', 'value': relay[0], 'qs_group': None})
    try:
        word_bits.append({'element': 'IPADDR', 'value': relay[3], 'qs_group': None})
    except IndexError:
        pass
    if pmu:
        pmu_id = {'element': 'PMSTN', 'value': relay[0], 'qs_group': None}
        word_bits.append(pmu_id)
    for row in settings[1:]:  # exclude headers
        if row[6] is not None:
            logic_class_list = str(row[6]).replace(' ', '').split(',')  # split logic class into array
            logic_class_list = [s.split('.')[0] for s in logic_class_list]
        else:
            logic_class_list = []
        if relay[2] is not None:
            logic_eval = str(relay[2]).split('.')[0] in logic_class_list
        else:
            logic_eval = False
        and_conditions = [row[5] is None, row[6] is None]
        or_conditions = [row[5] == relay[1], logic_eval]
        if all(and_conditions) or any(or_conditions):  # Class match
            if row[0] is not None:
                if isinstance(row[1], float) and row[float_index]:  # Round floats
                    formatted_string = "{:.2f}".format(row[1])  # To 2 decimal places
                    word_bits.append({'element': row[0], 'value': formatted_string, 'qs_group': row[8]})
                elif isinstance(row[1], float):
                    word_bits.append({'element': row[0], 'value': str(int(row[1])), 'qs_group': row[8]})
                else:
                    word_bits.append({'element': row[0], 'value': row[1], 'qs_group': row[8]})
    return word_bits


def gen_settings(xl_path, template_path, output_path, workbook_params):
    """Updates the rdb text based template which can be imported in QuickSet

    Args:
        xl_path (str): Path to the Excel workbook containing settings
        template_path (str): Path to the RDB template directory
        output_path (str): Path to the output directory where settings will be generated
        workbook_params (dict): {sheet_name, class_table, settings_table}
    """

    series_400 = ['XFMR_487E', 'CAP_487V']
    meters = ['Meter_735']
    dpac = ['DPAC_2440']
    try:
        app = xw.App(visible=False)
        wb = app.books.open(xl_path)
        sheet = wb.sheets[workbook_params['sheet_name']]

        # Get relay class info and create output directories
        relay_class = sheet.tables[workbook_params['class_table']].range.value
        relay_class = [item for item in relay_class if item[0] is not None]  # remove blank lines
        output_paths = []
        for relay in relay_class[1:]:
            if relay[0] is not None:
                new_dir = os.path.join(output_path, str(relay[0]))
                shutil.copytree(template_path, new_dir)
                output_paths.append(new_dir)

        settings = sheet.tables[workbook_params['settings_table']].range.value

        for i, relay in enumerate(relay_class[1:]):
            if relay[0] is not None:
                if workbook_params['sheet_name'] in meters:
                    word_bits = get_wordbits(relay, settings, mtr=True)
                elif workbook_params['sheet_name'] in dpac:
                    word_bits = get_wordbits(relay, settings, dpac=True)
                else:
                    word_bits = get_wordbits(relay, settings)
                # Generate RDB .txt file
                os.chdir(output_paths[i])
                if workbook_params['sheet_name'] in series_400:
                    update_template_400(word_bits)
                else:
                    update_template(word_bits)

                print(relay[0] + ' settings complete...')

    # Close workbook and quit app
    finally:
        wb.close()
        app.quit()


def update_template(word_bits):
    """Updates an SEL rdb template folder structure comprised of .txt files

    Args:
        word_bits (list): list comprised of dictionaries {element:val, value:val, qs_group:val}
        """

    file_names = os.listdir('.')
    # read file
    for file in file_names:
        found = []
        # read all .txt files
        if file.lower().endswith('.txt'):
            # isolate string segment after '_' and before '.TXT'
            settings_group = file.split("_")[1].split('.')[0]  # 'SET_1.TXT' -> 1 = quickset settings group
            file_handle = open(file, 'r')
            content = file_handle.readlines()
            file_handle.close()
            # traverse all excel variables and replace .txt line if there is a match
            for element in word_bits:
                for line in content:
                    if line.startswith(element['element'] + ',') and (settings_group == element['qs_group'] or
                                                                      element['qs_group'] is None):
                        index = content.index(line)
                        content[index] = element['element'] + ',"' + str(element['value']) + '"' + '\x1c\n'
                        found.append(index)
                        break

            if settings_group in ['D1']:
                for line in content:
                    index = content.index(line)
                    if index not in found and (len(line.split(',')) > 1):
                        temp = content[index].split(',')[0]
                        content[index] = temp + ',"NA"\x1c\n'

            # write new content to file
            file_handle = open(file, 'w', encoding="ascii")
            for line in content:
                file_handle.write(line)
            file_handle.close()


if __name__ == '__main__':
    # Example usage
    xl_path = r"C:\Users\laerps\OneDrive - Westwood Active Directory\Desktop\351S\Settings.xlsx"
    template_path = r"C:\Users\laerps\OneDrive - Westwood Active Directory\Desktop\351S\487V Template"
    output_path = r"C:\Users\laerps\OneDrive - Westwood Active Directory\Desktop\351S"
    gen_settings(xl_path, template_path, output_path, {'sheet_name': "CAP_487V", 'class_table': 'class_487V',
                                                       'settings_table': "settings_487V"})
