import os
import shutil
import xlwings as xw


def update_template_351S(word_bits):
    """Updates an SEL rdb template folder structure comprised of .txt files

    Args:
        word_bits (): list comprised of dictionaries {element:val, value:val, qs_group:val}
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
                        content[index] = element['element'] + ',"' + str(element['value']) + '"' + '\n'
                        found.append(index)
                        break

            if settings_group in ['D1']:
                for line in content:
                    index = content.index(line)
                    if index not in found and (len(line.split(',')) > 1):
                        temp = content[index].split(',')[0]
                        content[index] = temp + ',"NA"\n'

            # write new content to file
            file_handle = open(file, 'w')
            for line in content:
                file_handle.write(line)
            file_handle.close()


def get_wordbits(relay, settings):
    """Extracts word bits from settings table

    Args:
        relay (list): relay class definition including RID, IP, Settings Class, Logic Class, etc.
        settings (list): list including word bits and their associated values and properties
        """

    float_index = settings[0].index('Float')
    word_bits = [
        {'element': 'RID', 'value': relay[0], 'qs_group': None},
        {'element': 'IPADDR', 'value': relay[3], 'qs_group': None}
    ]
    for row in settings[1:]:  # exclude headers
        and_conditions = [row[5] is None, row[6] is None]
        or_conditions = [row[5] == relay[1], row[6] == relay[2]]
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


def gen_settings_351S(xl_path, template_path, output_path):
    """Updates the rdb text based template which can be imported in Quickset
    
    Args:
        xl_path (str): Path to the Excel workbook containing settings
        template_path (str): Path to the RDB template directory
        output_path (str): Path to the output directory where settings will be generated
    """

    try:
        app = xw.App(visible=False)
        wb = app.books.open(xl_path)
        sheet = wb.sheets['Feeder_351S']

        # Get relay class info and create output directories
        relay_class = sheet.tables['class_351S'].range.value
        relay_class = [item for item in relay_class if item[0] is not None]  # remove blank lines
        output_paths = []
        for relay in relay_class[1:]:
            if relay[0] is not None:
                new_dir = os.path.join(output_path, str(relay[0]))
                shutil.copytree(template_path, new_dir)
                output_paths.append(new_dir)

        settings = sheet.tables['settings_351S'].range.value
        get_wordbits(relay_class, settings)

        for i, relay in enumerate(relay_class[1:]):
            if relay[0] is not None:
                word_bits = get_wordbits(relay, settings)
                # Generate RDB .txt file
                os.chdir(output_paths[i])
                update_template_351S(word_bits)

                print(relay[0] + ' settings complete...')

    # Close workbook and quit app
    finally:
        wb.close()
        app.quit()


def gen_settings_HV351S(xl_path, template_path, output_path):
    """Updates the rdb text based template which can be imported in QuickSet

    Args:
        xl_path (str): Path to the Excel workbook containing settings
        template_path (str): Path to the RDB template directory
        output_path (str): Path to the output directory where settings will be generated
    """

    try:
        app = xw.App(visible=False)
        wb = app.books.open(xl_path)
        sheet = wb.sheets['HV_351S']

        # Get relay class info and create output directories
        relay_class = sheet.tables['class_HV351S'].range.value
        relay_class = [item for item in relay_class if item[0] is not None]  # remove blank lines
        output_paths = []
        for relay in relay_class[1:]:
            if relay[0] is not None:
                new_dir = os.path.join(output_path, str(relay[0]))
                shutil.copytree(template_path, new_dir)
                output_paths.append(new_dir)

        settings = sheet.tables['settings_HV351S'].range.value
        get_wordbits(relay_class, settings)

        for i, relay in enumerate(relay_class[1:]):
            if relay[0] is not None:
                word_bits = get_wordbits(relay, settings)
                # Generate RDB .txt file
                os.chdir(output_paths[i])
                update_template_351S(word_bits)

                print(relay[0] + ' settings complete...')

    # Close workbook and quit app
    finally:
        wb.close()
        app.quit()

if __name__ == '__main__':
    # Example usage
    xl_path = r"C:\Users\laerps\OneDrive - Westwood Active Directory\Desktop\351S\Settings.xlsx"
    template_path = r"C:\Users\laerps\OneDrive - Westwood Active Directory\Desktop\351S\test template"
    output_path = r"C:\Users\laerps\OneDrive - Westwood Active Directory\Desktop\351S"
    gen_settings_HV351S(xl_path, template_path, output_path)
