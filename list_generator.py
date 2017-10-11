from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.copier import WorksheetCopy
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet import Worksheet
from saar_teams import *
import datetime


class ExcelGenerator(object):

    def __init__(self, file_prefix):
        self._wb = None
        self._file_prefix = file_prefix

    def generate(self, team_list):
        raise NotImplementedError()

    def write(self):
        now = datetime.datetime.now()
        self._wb.save("{}_{}.xlsx".format(self._file_prefix, now.strftime("%Y-%m-%d")))

    def close(self):
        self._wb.close()


class RefereeFormsGenerator(ExcelGenerator):

    def __init__(self):
        super(RefereeFormsGenerator, self).__init__("Wertungsbogen")

    def generate(self, team_list: SaarTeamList):

        self._wb = load_workbook(filename="./Dokumente/Wertungsbogen_Master.xlsx")
        master_ws = self._wb.active
        version_string = master_ws["F2"].value
        check_version(version_string)

        i = 0
        for team in team_list:
            assert isinstance(team, SaarTeam)
            for apparatus_f, apparatus_m in [("Boden", "Boden"), ("Sprung", "Sprung"), ("Stufenbarren", "Reck"), ("Balken", "Barren")]:
                title = "{}_{}_{}".format(team.name[:10], apparatus_f[:2], apparatus_m[:2])
                print(title)

                ws = self._wb.create_sheet(title=title)
                ws_copy = WorksheetCopy(master_ws, ws)
                ws_copy.copy_worksheet()

                # set smaller page margins
                assert isinstance(ws, Worksheet)
                ws.page_margins = PageMargins(.2, .2, .75, .75, .314, .314)

                # set contents
                ws["F1"] = apparatus_f
                ws["F18"] = apparatus_m
                ws["A2"] = team.name
                ws["A19"] = team.name

                offset = 4
                for num, gymnast in enumerate(team.get_gymnasts(SaarGymnast.FEMALE)):
                    row = offset + num
                    ws["A{}".format(row)] = num + 1
                    ws["B{}".format(row)] = gymnast.name
                    ws["C{}".format(row)] = gymnast.surname

                offset = 21
                for num, gymnast in enumerate(team.get_gymnasts(SaarGymnast.MALE)):
                    row = offset + num
                    ws["A{}".format(row)] = num + 1
                    ws["B{}".format(row)] = gymnast.name
                    ws["C{}".format(row)] = gymnast.surname

                i += 1


