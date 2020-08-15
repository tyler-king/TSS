import os
import shutil
from oletools.olevba3 import VBA_Parser
from pathlib import Path

EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)
KEEP_NAME = False  # Set this to True if you would like to keep "Attribute VB_Name"


def parse(workbook_path):
    vba_path = 'src.vba'
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []
    workbook_name = Path(workbook_path).stem
	
    for _, _, filename, content in vba_modules:
        decoded_content = content # .decode('latin-1')
        lines = []
        if '\r\n' in decoded_content:
            lines = decoded_content.split('\r\n')
        else:
            lines = decoded_content.split('\n')
        if lines:
            content = []
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    if not os.path.exists(os.path.join(vba_path, workbook_name)):
                        os.makedirs(os.path.join(vba_path, workbook_name))
                    with open(os.path.join(vba_path, workbook_name, filename), 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content))


if __name__ == '__main__':
    for root, dirs, files in os.walk('.'):
        for f in dirs:
            if f.endswith('.vba'):
                shutil.rmtree(os.path.join(root, f))

        for f in files:
            if f.endswith(EXCEL_FILE_EXTENSIONS):
                parse(os.path.join(root, f))
