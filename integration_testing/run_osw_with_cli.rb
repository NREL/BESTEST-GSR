require 'fileutils'

path_datapoints = "workflow"

# loop through resoruce files
jobs = []
results_directories = Dir.glob("#{path_datapoints}/*")
results_directories.each do |directory|
	puts "queueing up #{directory}"
	test_dir = "#{directory}/data_point.osw"
	string = "openstudio run -w '#{test_dir}'"
	if not File.file?(test_dir)
	  puts "data_point.osw not found for #{directory}"
	  next
	end

	# system(string)
	jobs << string

end

# run the jobs
# if gem parallel isn't isntalled then comment out this coudl and use system(string) to run one job at a time
require 'parallel'
num_parallel = 12 #6
Parallel.each(jobs, in_threads: num_parallel) do |job|
  puts job
  system(job)
end