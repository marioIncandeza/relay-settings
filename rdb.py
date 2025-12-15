import os
import shutil
import xlwings as xw


# --- Configuration Constants ---
SERIES_400_DEVICES = ['XFMR_487E', 'CAP_487V', 'Line_411L']
METER_DEVICES = ['MTR_735']
DPAC_DEVICES = ['DPAC_2440']

# Mapping logic for different device families
# This replaces the need for two separate 'update_template' functions
DEVICE_CONFIGS = {
    'SERIES_400': {
        'clear_val': '""',  # 400 series uses empty string for clears
        'clear_groups': [
            'D1', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6',
            'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10'
        ],
        'process_f1': True
    },
    'STANDARD': {
        'clear_val': '"NA"',  # Standard uses "NA"
        'clear_groups': ['D1'],
        'process_f1': False
    }
}


# --- Shared Configuration data (Moved from main.py) ---
RELAY_CONFIG = {
    'feeder': {'label': 'Feeder 351S', 'params': {'sheet_name': 'FDR_351S', 'class_table': 'class_351S',
                                                  'settings_table': 'settings_351S'}},
    'hv': {'label': 'HV 351S', 'params': {'sheet_name': 'HV_351S', 'class_table': 'class_HV351S',
                                          'settings_table': 'settings_HV351S'}},
    'xfmr_487E': {'label': 'XFMR 487E', 'params': {'sheet_name': 'XFMR_487E', 'class_table': 'class_487E',
                                                   'settings_table': 'settings_487E'}},
    'cap_487V': {'label': 'CAP 487V', 'params': {'sheet_name': 'CAP_487V', 'class_table': 'class_487V',
                                                 'settings_table': 'settings_487V'}},
    'bus_587Z': {'label': 'BUS 587Z', 'params': {'sheet_name': 'BUS_587Z', 'class_table': 'class_587Z',
                                                 'settings_table': 'settings_587Z'}},
    'mtr_735': {'label': 'MTR 735', 'params': {'sheet_name': 'MTR_735', 'class_table': 'class_735',
                                               'settings_table': 'settings_735'}},
    'dpac_2440': {'label': 'DPAC 2440', 'params': {'sheet_name': 'DPAC_2440', 'class_table': 'class_2440',
                                                   'settings_table': 'settings_2440'}},
    'xfmr_787': {'label': 'XFMR 787', 'params': {'sheet_name': 'XFMR_787', 'class_table': 'class_787',
                                                 'settings_table': 'settings_787'}},
    'line_411L': {'label': 'LINE 411L', 'params': {'sheet_name': 'Line_411L', 'class_table': 'class_411L',
                                                   'settings_table': 'settings_411L'}}
}

# Shared definition for 351S style groups
_common_351_groups = {
    "labels": [f"Group {i}" for i in range(1, 7)] + [f"Logic {i}" for i in range(1, 7)],
    "shorthand": {**{f"Group {i}": str(i) for i in range(1, 7)}, **{f"Logic {i}": f"L{i}" for i in range(1, 7)}}
}

_common_400_groups = {
    "labels": [f"Set {i}" for i in range(1, 7)] + [f"Protection Logic {i}" for i in range(1, 7)],
    "shorthand": {**{f"Set {i}": f"S{i}" for i in range(1, 7)},
                  **{f"Protection Logic {i}": f"L{i}" for i in range(1, 7)}}
}

_common_787_groups = {
    "labels": [f"Set {i}" for i in range(1, 5)] + [f"Logic {i}" for i in range(1, 5)],
    "shorthand": {**{f"Set {i}": str(i) for i in range(1, 5)}, **{f"Logic {i}": f"L{i}" for i in range(1, 5)}}
}

RELAY_REGION_METADATA = {
    "feeder": _common_351_groups,
    "hv": _common_351_groups,
    "xfmr_487E": _common_400_groups,
    "cap_487V": _common_400_groups,
    "xfmr_787": _common_787_groups,
    "line_411L": _common_400_groups
}


def gen_settings(xl_path, template_path, output_path, workbook_params, excluded_regions=None, include_comments=True):
    """
    Main driver function to generate settings.
    Added include_comments parameter.
    """
    if excluded_regions is None:
        excluded_regions = []

    sheet_name = workbook_params['sheet_name']

    if sheet_name in SERIES_400_DEVICES:
        config = DEVICE_CONFIGS['SERIES_400']
    else:
        config = DEVICE_CONFIGS['STANDARD']

    is_mtr = sheet_name in METER_DEVICES
    is_dpac = sheet_name in DPAC_DEVICES

    app = xw.App(visible=False)
    try:
        wb = app.books.open(xl_path)
        sheet = wb.sheets[sheet_name]

        relay_class_rng = sheet.tables[workbook_params['class_table']].range.value
        settings_rng = sheet.tables[workbook_params['settings_table']].range.value
        relay_class = [item for item in relay_class_rng if item[0] is not None]

        # 1. Create Directories
        output_dirs = []
        valid_relays = []

        for relay in relay_class[1:]:
            if relay[0] is not None:
                new_dir = os.path.join(output_path, str(relay[0]))
                if os.path.exists(new_dir):
                    shutil.rmtree(new_dir)
                shutil.copytree(template_path, new_dir)
                output_dirs.append(new_dir)
                valid_relays.append(relay)

        # 2. Process Settings
        for i, relay in enumerate(valid_relays):
            print(f"Processing {relay[0]}...")

            # Pass the comment toggle down to extraction logic
            word_bits = get_wordbits(relay, settings_rng, mtr=is_mtr, dpac=is_dpac, include_comments=include_comments)

            process_rdb_files(output_dirs[i], word_bits, excluded_regions, config)

            print(f"{relay[0]} settings complete.")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        raise
    finally:
        wb.close()
        app.quit()


def get_wordbits(relay, settings, pmu=True, mtr=False, dpac=False, include_comments=True):
    """Extracts word bits from settings table

    Args:
        relay (list): relay class definition including RID, IP, Settings Class, Logic Class, etc.
        settings (list): list including word bits and their associated values and properties
        pmu (bool): include PMU station name
        """

    def get_cmt(text):
        return text if include_comments else ""

    float_index = settings[0].index('Float')
    word_bits = []
    if mtr:
        word_bits.append({'element': 'MID', 'value': relay[0], 'qs_group': None, 'comment': get_cmt('Meter ID')})
    elif dpac:
        word_bits.append({'element': 'DID', 'value': relay[0], 'qs_group': None, 'comment': get_cmt('Device ID')})
    else:
        word_bits.append({'element': 'RID', 'value': relay[0], 'qs_group': None, 'comment': get_cmt('Relay ID')})
    try:
        word_bits.append({'element': 'IPADDR', 'value': relay[3], 'qs_group': None, 'comment': get_cmt('IP Address')})
    except IndexError:
        pass
    if pmu:
        pmu_id = {'element': 'PMSTN', 'value': relay[0], 'qs_group': None, 'comment': get_cmt('Phasor ID')}
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
                    word_bits.append(
                        {'element': row[0], 'value': formatted_string, 'qs_group': row[8], 'comment': get_cmt(row[2])})
                elif isinstance(row[1], float):
                    word_bits.append(
                        {'element': row[0], 'value': str(int(row[1])), 'qs_group': row[8], 'comment': get_cmt(row[2])})
                else:
                    word_bits.append({'element': row[0], 'value': row[1], 'qs_group': row[8], 'comment': get_cmt(row[2])})
    return word_bits


def process_rdb_files(target_dir, word_bits, excluded_regions, config):
    """
    Unified function to process RDB text files.
    Replaces both update_template and update_template_400.
    """
    if excluded_regions is None:
        excluded_regions = []

    # Optimize: Create a lookup dictionary for word bits to avoid nested loops.
    # Key: Element Name, Value: Bit Data
    # Note: If duplicate elements exist across different groups, this logic holds
    # because we check qs_group match inside the loop.
    wb_lookup = {wb['element']: wb for wb in word_bits}

    for file_name in os.listdir(target_dir):
        if not file_name.lower().endswith('.txt'):
            continue

        # Parse Group from filename (e.g., 'SET_1.TXT' -> '1')
        try:
            parts = file_name.split('_')
            if len(parts) > 1:
                settings_group = parts[1].split('.')[0]
            else:
                settings_group = "UNKNOWN"
        except IndexError:
            continue

        if settings_group in excluded_regions:
            continue

        file_path = os.path.join(target_dir, file_name)

        # Read content
        with open(file_path, 'r') as f:
            lines = f.readlines()

        new_lines = []
        found_indices = set()

        # PASS 1: Update values from Word Bits
        for idx, line in enumerate(lines):
            line_parts = line.split(',')
            if not line_parts:
                new_lines.append(line)
                continue

            element_key = line_parts[0]

            # Check if this element exists in our Excel data
            if element_key in wb_lookup:
                wb = wb_lookup[element_key]

                # Check Group Constraint
                if wb['qs_group'] is None or str(wb['qs_group']) == settings_group:
                    # Only update if we have a value
                    if wb['value']:
                        # Construct SEL RDB format: ELEMENT,"VALUE"<0x1c>COMMENT
                        new_line = f'{wb["element"]},"{wb["value"]}"\x1c{wb["comment"]}\n'
                        new_lines.append(new_line)
                        found_indices.add(idx)
                        continue

            new_lines.append(line)

        # PASS 2: Clear Logic (D1, L1... or F1 specific handling)
        final_lines = []

        # Determine if this file needs clearing logic
        needs_clearing = settings_group in config['clear_groups']
        is_f1 = (settings_group == 'F1') and config['process_f1']

        if needs_clearing or is_f1:
            for idx, line in enumerate(new_lines):
                # Skip lines we just updated
                if idx in found_indices:
                    final_lines.append(line)
                    continue

                line_parts = line.split(',')
                if len(line_parts) <= 1:
                    final_lines.append(line)
                    continue

                element_key = line_parts[0]

                # Clear Logic for specified groups
                if needs_clearing:
                    # Set value to configured clear value ("" or "NA")
                    # Note: \x1c is the field separator in SEL RDB
                    cleared_line = f'{element_key},{config["clear_val"]}\x1c\n'
                    final_lines.append(cleared_line)

                # F1 Specific Logic (DP_NAM/DP_SIZE)
                elif is_f1 and (line.startswith('DP_NAM') or line.startswith('DP_SIZE')):
                    # Check if we have a comment in the original lookup to preserve?
                    # Original code used the last 'element' loop variable, which was buggy.
                    # We will append a generic closure or empty comment.
                    cleared_line = f'{element_key},""\x1c\n'
                    final_lines.append(cleared_line)

                else:
                    final_lines.append(line)
        else:
            final_lines = new_lines

        # Write back to file

        with open(file_path, 'w', encoding='ascii') as f:
            f.writelines(final_lines)


def get_template_info(template_path):
    """
    Parses the [INFO] section of Misc/Cfg.txt in the template directory.
    Returns a dict of key-value pairs.
    """
    cfg_path = os.path.join(template_path, 'Misc', 'Cfg.txt')
    info_dict = {}

    if not os.path.exists(cfg_path):
        return info_dict

    try:
        with open(cfg_path, 'r') as f:
            lines = f.readlines()

        in_info_section = False
        for line in lines:
            line = line.strip()
            if line == '[INFO]':
                in_info_section = True
                continue
            if line.startswith('['): # Any other section
                in_info_section = False
            
            if in_info_section and '=' in line:
                key, value = line.split('=', 1)
                info_dict[key.strip()] = value.strip()
                
    except Exception as e:
        print(f"Error reading template info: {e}")

    return info_dict


if __name__ == '__main__':
    # Example usage
    xl_path_ex = r"C:\Users\Example\Desktop\351S\Settings.xlsx"
    template_path_ex = r"C:\Users\Example\Desktop\351S\487V Template"
    output_path_ex = r"C:\Users\Example\Desktop\351S"

    # Ensure params match your Excel table names exactly
    params = {
        'sheet_name': "CAP_487V",
        'class_table': 'class_487V',
        'settings_table': "settings_487V"
    }

    gen_settings(xl_path_ex, template_path_ex, output_path_ex, params)
