import datetime
import os


class ListGenerator(object):
    def __init__(self, path):
        self._path = path

    def _get_path(self, filename):
        return os.path.join(self._path, filename)

    def generate(self, team_list):
        raise NotImplementedError()

    def write(self):
        raise NotImplementedError()

    def close(self):
        raise NotImplementedError()


class ExcelGenerator(ListGenerator):
    def __init__(self, path, prefix):
        super().__init__(path)
        self._wb = None
        self._prefix = prefix

    def generate(self, team_list):
        raise NotImplementedError()

    def write(self):
        now = datetime.datetime.now()
        filename = "{}_{}.xlsx".format(self._prefix, now.strftime("%Y-%m-%d"))
        self._wb.save(super()._get_path(filename))

    def close(self):
        self._wb.close()



