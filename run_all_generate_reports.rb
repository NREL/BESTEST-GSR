require 'fileutils'

# this will run OSW for all BESTEST test cases, create a combined CSV, copy zip files, and generate updated Excel files.
# At a minimum there should be an argument to not copy zip files, could expose other options
# clean up workflow directory at end so excess files are not sent to repo.

Dir.chdir("integration_testing")

puts "generate test cases and running simulations"
load 'run_osw_with_cli.rb'

puts "creating combined summmary CSV file"
system("openstudio gather_results_generate_csv.rb")

# not setup to handle failed datapoints, will stop if any datapoint is missing expected files
puts "copy ZIP file"
system("openstudio rename_and_gather_bestest_zips.rb")

Dir.chdir("../results")

puts "generate envelope Excel Reports"
load 'bestest_populate_report.rb'

puts "generate Cooling Excel Reports"
load 'bestest_populate_report_CE.rb'

puts "generate Heating Excel Reports"
load 'bestest_populate_report_HE.rb'

Dir.chdir("../")

puts "cleaning up workflow directories"
workflow_directories = Dir.glob("integration_testing/workflow/*")
workflow_directories.each do |directory|
	next if directory.incldue?("workflow_resources")
	content =  Dir.glob("#{directory}/*")
	content.each do |file|
	  next if file.include?("data_point.osw")
	  next if file.include?("workflow.osw")
	  FileUtils.rm_r(file)
	end
end
