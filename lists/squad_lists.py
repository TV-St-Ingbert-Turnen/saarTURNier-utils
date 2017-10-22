from openpyxl import Workbook, load_workbook
from openpyxl.worksheet import Worksheet
from openpyxl.worksheet.copier import WorksheetCopy
from openpyxl.worksheet.page import PageMargins
import csv
from saar_teams import *
from constants import APPARATUS
from .list_generator import ListGenerator, ExcelGenerator
import os
from copy import copy

class CompetitionPlanGenerator(ExcelGenerator):
    def __init__(self, path, num_squads):
        super().__init__(path, "Wettkampfplan")
        self._num_squads = num_squads

    def generate(self, team_list: SaarTeamList):

        self._wb = load_workbook(filename=os.path.join(os.getcwd(), "templates/Wettkampfplan.xlsx"))

        ws = self._wb.active

        squads = team_list.get_squads(num_squads=self._num_squads)

        # generate simple squad list
        row_offset = 31
        for sq_id, sq in enumerate(squads):
            for team in sq:
                ws["A{}".format(row_offset)].value = sq_id + 1
                ws["B{}".format(row_offset)].value = team.name
                row_offset += 1

        # generate referee assignment table
        # one referee-set per squad (male/female); 4 are pre-defined
        if len(squads) == 2:
            # remove additional cols for referees
            for row in ws.iter_rows(min_row=19, max_row=25, min_col=4, max_col=5):
                for cell in row:
                    cell.value = None
                    cell.border = None
            # copy referee legend left
            for row in ws.iter_rows(min_row=20, max_row=25, min_col=6, max_col=6):
                for cell in row:
                    n_cell = ws.cell(row=cell.row, column=cell.col_idx-2)
                    n_cell.value = cell.value
                    n_cell.font = copy(cell.font)
                    cell.border = None
                    cell.value = None

        # generate competition plan
        num_rotations = 4
        for rotation in range(4):
            for j, sq in enumerate(squads):
                apparatus = divmod(rotation + j, num_rotations)[1]