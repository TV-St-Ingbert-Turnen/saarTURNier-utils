from openpyxl import Workbook, load_workbook
from openpyxl.worksheet import Worksheet
from openpyxl.worksheet.copier import WorksheetCopy
from openpyxl.worksheet.page import PageMargins
import csv
from saar_teams import *
from constants import APPARATUS
from .list_generator import ListGenerator, ExcelGenerator


class CompetitionPlanGenerator(ExcelGenerator):
    def __init__(self, path, num_squads):
        super().__init__(path, "Wettkampfplan")
        self._num_squads = num_squads

    def generate(self, team_list: SaarTeamList):
        self._wb = Workbook()

        ws = self._wb.active

        squads = team_list.get_squads(num_squads=self._num_squads)

        # generate simple squad list
        for sq in squads:
            # squad id (starting from 1)
            # squad teams
            pass

        # generate referee assignment table
        for i in range(len(squads)):
            # one referee-set per squad (male/female)

            # apparatus, D and up to 2 E referees for both, male and female (also per squad)
            pass

        # generate competition plan
        num_rotations = 4
        for rotation in range(4):
            for j, sq in enumerate(squads):
                apparatus = divmod(rotation + j, num_rotations)[1]


        i = 0
        for team in team_list:
            assert isinstance(team, SaarTeam)
            for apparatus_f, apparatus_m in APPARATUS:
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

        self._wb.remove(self._wb["Master"])