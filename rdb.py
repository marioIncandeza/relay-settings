import os
import shutil
import xlwings as xw


def update_template_351S(word_bits):
    """This function updates an SEL rdb template"""
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


def gen_settings_351S(xl_path, template_path, output_path):
    """This function updates the rdb text based template which can be imported in Quickset
    
    Args:
        xl_path (str): Path to the Excel workbook containing settings
        template_path (str): Path to the RDB template directory
        output_path (str): Path to the output directory where settings will be generated
    """

    try:
        app = xw.App(visible=False)
        wb = app.books.open(xl_path)
        sheet = wb.sheets['Feeder_351S']

        # Get RDB Variables and create output directories
        relay_class = sheet.tables['class_351S'].range.value
        output_paths = []
        for relay in relay_class[1:]:
            new_dir = os.path.join(output_path, str(relay[0]))
            shutil.copytree(template_path, new_dir)
            output_paths.append(new_dir)

        settings = sheet.tables['settings_351S'].range.value
        float_index = settings[0].index('Float')
        settings.pop(0)  # Remove headers

        for i, relay in enumerate(relay_class[1:]):
            word_bits = [
                {'element': 'RID', 'value': relay[0], 'qs_group': None},
                {'element': 'IPADDR', 'value': relay[3], 'qs_group': None}
            ]
            for row in settings:
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
    xl_path = "path/to/settings.xlsx"
    template_path = "path/to/template"
    output_path = "path/to/output"
    gen_settings_351S(xl_path, template_path, output_path)
