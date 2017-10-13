from saar_teams import SaarTeamList
from list_generator import RefereeFormsGenerator, ScoreSystemCsvGenerator

if __name__ == '__main__':
    team_list = SaarTeamList("Sample/Teilnehmerliste.xlsx")
    teams = team_list.teams

    #gen = RefereeFormsGenerator()
    #gen.generate(team_list)
    #gen.write()
    #gen.close()

    gen = ScoreSystemCsvGenerator()
    gen.generate(team_list)
    gen.write()
    gen.close()

    pass
