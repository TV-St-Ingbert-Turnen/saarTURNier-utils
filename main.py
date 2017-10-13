from saar_teams import SaarTeamList
from list_generator import RefereeFormsGenerator, ScoreSystemCsvGenerator, RefereeCsvGenerator
import os

if __name__ == '__main__':
    # input test
    default_input_path = "input_files"
    team_list = SaarTeamList(os.path.join(os.getcwd(), default_input_path, "Teilnehmerliste.xlsx"))
    teams = team_list.teams


    # output test
    default_output_path = "generated_lists"

    gen = RefereeFormsGenerator(default_output_path)
    gen.generate(team_list)
    gen.write()
    gen.close()

    gen = ScoreSystemCsvGenerator(default_output_path)
    gen.generate(team_list)
    gen.write()
    gen.close()

    gen = RefereeCsvGenerator(default_output_path)
    gen.generate(team_list)
    gen.write()
    gen.close()

    pass
