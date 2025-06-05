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


def update_template_487E(word_bits):
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
                        if element['value']:
                            content[index] = element['element'] + ',"' + str(element['value']) + '"' + '\x1c\n'
                        found.append(index)
                        break

            if settings_group in ['D1']:
                for line in content:
                    index = content.index(line)
                    if index not in found and (len(line.split(',')) > 1):
                        temp = content[index].split(',')[0]
                        content[index] = temp + ',""\x1c\n'

            # write new content to file
            file_handle = open(file, 'w', encoding="ascii")
            for line in content:
                file_handle.write(line)
            file_handle.close()


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