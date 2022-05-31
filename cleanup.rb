require 'fileutils'

puts "cleaning up workflow directories"
workflow_directories = Dir.glob("integration_testing/workflow/*")
workflow_directories.each do |directory|
	next if directory.include?("workflow_resources")
	content =  Dir.glob("#{directory}/*")
	content.each do |file|
	  next if file.include?("data_point.osw")
	  next if file.include?("workflow.osw")
	  FileUtils.rm_rf(file)
	end
end