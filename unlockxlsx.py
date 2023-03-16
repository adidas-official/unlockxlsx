import shutil
from shutil import copyfile
from pathlib import Path
import zipfile
import re


class Locksmith:
    def __init__(self, xlsx_file, sheet_nums):
        self.xlsx = xlsx_file
        self.dest_dir = './tempdir'
        self.sheet_nums = [str(sheet_num) for sheet_num in sheet_nums]

    def make_a_dest(self):
        if not (n := Path(self.dest_dir)).exists():
            n.mkdir()

    def make_a_copy_zip(self):
        self.make_a_dest()
        xlsx = Path(self.dest_dir) / Path(self.xlsx)
        zip_name = Path(self.dest_dir) / (xlsx.stem + '.zip')
        if not xlsx.exists():
            copyfile(self.xlsx, zip_name)

    def remove_protection(self):
        self.make_a_copy_zip()
        with zipfile.ZipFile('tempdir/' + self.xlsx.replace('xlsx', 'zip'), mode='a') as zf:

            for sheet in self.sheet_nums:
                sheet_path = f'xl/worksheets/sheet{sheet}.xml'
                content = zf.read(sheet_path).decode('cp1250')
                pattern = re.compile(r'<sheetProtection.*?/>')
                result = re.sub(pattern, '', content)

                zf.writestr(sheet_path, result)

    def make_an_xslx(self):
        self.remove_protection()
        zip_dest = Path(self.dest_dir) / (Path(self.xlsx).stem + '.zip')
        shutil.copyfile(zip_dest, Path('.') / ('unlocked_' + self.xlsx))
        shutil.rmtree(self.dest_dir)
