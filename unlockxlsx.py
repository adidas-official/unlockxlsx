import shutil
from shutil import copyfile
from pathlib import Path
import zipfile
import re


class Locksmith:
    def __init__(self, xlsx_file, dest_dir, sheet_num):
        self.xlsx = xlsx_file
        self.dest_dir = dest_dir
        self.sheet_num = str(sheet_num)

    def make_a_dest(self):
        if not (n := Path(self.dest_dir)).exists():
            n.mkdir()

    def make_a_copy_zip(self):
        self.make_a_dest()
        xlsx = Path(self.dest_dir) / Path(self.xlsx)
        zip_name = Path(self.dest_dir) / (xlsx.stem + '.zip')
        if not xlsx.exists():
            copyfile(self.xlsx, zip_name)

    def unzip_file(self):
        self.make_a_copy_zip()
        zip_name = Path(self.dest_dir) / (Path(self.xlsx).stem + '.zip')
        with zipfile.ZipFile(zip_name, 'r') as zip_ref:
            zip_ref.extractall(self.dest_dir)

    def remove_protection(self):
        self.unzip_file()
        sheet = Path(self.dest_dir) / 'xl' / 'worksheets' / f'sheet{self.sheet_num}.xml'
        with open(sheet, 'r+') as f:
            content = f.read()
            pattern = re.compile(r'<sheetProtection.*?/>')
            result = re.sub(pattern, '', content)
            f.seek(0)
            f.write(result)

    def add_to_archive(self):
        self.remove_protection()
        sheet = Path(self.dest_dir) / 'xl' / 'worksheets' / f'sheet{self.sheet_num}.xml'
        zip_dest = Path(self.dest_dir) / (Path(self.xlsx).stem + '.zip')
        with zipfile.ZipFile(zip_dest, 'a') as zipf:
            zipf.write(sheet, '/xl/worksheets/sheet1.xml')

    def make_an_xslx(self):
        self.add_to_archive()
        zip_dest = Path(self.dest_dir) / (Path(self.xlsx).stem + '.zip')
        shutil.copyfile(zip_dest, Path('.') / ('unlocked_' + self.xlsx))
        shutil.rmtree(self.dest_dir)


locksmith = Locksmith('temp-up.xlsx', 'testdir', 1)
locksmith.make_an_xslx()
