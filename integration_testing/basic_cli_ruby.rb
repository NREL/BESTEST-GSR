# get absolute path to this script
script_path = File.dirname(__FILE__)

# setup absolute path so this script can be called from any directory
workflow_path = script_path + "/workflow/600EN_from_osm/workflow.osw"

# setup stirng for CLI
cli_string = "openstudio run -w " + workflow_path
print(cli_string)

# call OpenStudio CLI from Ruby for example OSW
system(cli_string)