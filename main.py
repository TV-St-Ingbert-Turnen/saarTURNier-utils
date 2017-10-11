from saar_teams import SaarTeamList
from list_generator import RefereeFormsGenerator

if __name__ == '__main__':
    team_list = SaarTeamList("../2017/Teilnehmerliste.xlsx")
    teams = team_list.teams

    gen = RefereeFormsGenerator()
    gen.generate(team_list)
    gen.write()

    pass
