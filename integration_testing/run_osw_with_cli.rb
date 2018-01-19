require 'fileutils'

path_datapoints = "workflow"

# loop through resoruce files
results_directories = Dir.glob("#{path_datapoints}/*")
results_directories.each do |directory|
	puts "runing #{directory}"
	test_dir = "#{directory}/data_point.osw"
	string = "openstudio run -w '#{test_dir}'"
	if not File.file?(test_dir)
	  puts "data_point.osw not found for #{directory}"
	  next
	end
	system(string)
end