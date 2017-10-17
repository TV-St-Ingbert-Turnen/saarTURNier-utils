from .list_generator import ListGenerator
import csv
from saar_teams import *


class ScoreSystemCsvGenerator(ListGenerator):
    def __init__(self, path):
        super().__init__(path)
        self._teams = []
        self._participants = []
        # participants.csv: "Vorname Name"; "[w|m]"; "tid"
        # teams.csv: "tid";"Name"

    def generate(self, team_list):
        for tid, team in enumerate(team_list):
            assert isinstance(team, SaarTeam)
            self._teams.append([tid + 1, team.name])
            for gymnast in team.gymnasts:
                g_name = "{} {}".format(gymnast.name, gymnast.surname)
                g_gender = 'm' if gymnast.gender == SaarGymnast.MALE else 'w'
                self._participants.append([g_name, g_gender, tid + 1])

    def write(self):
        with open(super()._get_path('teams.csv'), 'w', newline='\n', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_ALL)
            writer.writerows(self._teams)
        with open(super()._get_path('participants.csv'), 'w', newline='\n', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_ALL)
            writer.writerows(self._participants)

    def close(self):
        pass
