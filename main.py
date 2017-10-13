from saar_teams import SaarTeamList
from list_generator import RefereeFormsGenerator, ScoreSystemCsvGenerator, RefereeCsvGenerator

if __name__ == '__main__':
    team_list = SaarTeamList("input_files/Teilnehmerliste.xlsx")
    teams = team_list.teams

    base_path = "generated_lists"

    gen = RefereeFormsGenerator(base_path)
    gen.generate(team_list)
    gen.write()
    gen.close()

    gen = ScoreSystemCsvGenerator(base_path)
    gen.generate(team_list)
    gen.write()
    gen.close()

    gen = RefereeCsvGenerator(base_path)
    gen.generate(team_list)
    gen.write()
    gen.close()

    pass
