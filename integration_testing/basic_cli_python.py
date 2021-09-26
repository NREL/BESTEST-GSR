# get absolute path to this script
import os
script_path = os.path.dirname(__file__)

# setup absolute path so this script can be called from any directory
workflow_path = script_path + "/workflow/600EN_from_osm/workflow.osw"

# setup stirng for CLI
cli_string = "openstudio run -w " + workflow_path
print(cli_string)

# call OpenStudio CLI from Python for example OSW (should change ot use subprocess instead)	
os.system(cli_string)
