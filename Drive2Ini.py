
import sys, os

sys.path.insert(0, os.path.realpath("./Ini2Csv"))

from loader import Conf
from synchronize import Synchronize

class Ini(object):
    def __init__(self, conf_name):
        self._cols = []
        self._conf = Conf(conf_name)
        self._datas = self._conf.getAll()
        self._rows = []
        self._prepare_columns()

    def _prepare_columns(self):
        for k, v in self._datas.items():
            for col in v:
                if col not in self._cols:
                    self._cols.append(col)

    def _generate_header(self, first="slug"):
        row = []
        row.append(first)
        for key in self._cols:
            row.append(key)
        return row

    def _generate_rows(self):
        for k, v in self._datas.items():
            row = []
            row.append(k)
            for key in self._cols:
                row.append((", ".join(v[key]) if isinstance(v[key], list) else v[key] ) 
                           if key in v else "")
            self._rows.append(row)

    def export(self):
        self._rows = []
        self._generate_rows()
        self._rows.insert(0, self._generate_header())
        return self._rows

class Drive2Ini(object):
    def __init__(self, filename):
        self._filename = filename
        self._s = Synchronize()

    def toDrive(self, fast=True):
        o = Ini(self._filename)
        if fast:
            return self._s.fastLoad(o.export())        
        return self._s.load(o.export())

    def toFile(self):
        datas = self._s.export()
        conf = Conf(self._filename)
        for k, v in datas.items():
            conf.removeSection(k)
            for key, value in v.items():
                if not value or len(str(value)) == 0 or key == "slug":
                    del v[key]
                if value == "TRUE":
                    v[key] = "true"
            if len(v) > 0:
                conf.setSection(k, v)
        conf.save()

if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.stderr.write("USAGE: %s toDrive|toFile file.ini\n" % sys.argv[0])
        sys.exit(1)
    obj = Drive2Ini(sys.argv[2])
    if sys.argv[1] == "toDrive":
        obj.toDrive()
    elif sys.argv[1] == "toFile":
        obj.toFile()
