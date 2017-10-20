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
        self._wb = load_workbook(filename="../templates/Wettkampfplan.xlsx")

        ws = self._wb.active

        squads = team_list.get_squads(num_squads=self._num_squads)

        # generate simple squad list
        row_offset = 31
        for sq_id, sq in enumerate(squads):
            for team in sq:
                ws["A{}".format(row_offset)] = sq_id + 1
                ws["B{}".format(row_offset)] = team.name
                row_offset += 1

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