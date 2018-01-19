require 'fileutils'
require 'openstudio'

path_datapoints = "workflow"


# loop through resoruce files
results_directories = Dir.glob("#{path_datapoints}/*")
results_directories.each do |results_directory|

	# load the test model
	translator = OpenStudio::OSVersion::VersionTranslator.new
	path = OpenStudio::Path.new("#{File.dirname(__FILE__)}/#{results_directory}/run/in.osm")
	model = translator.loadModel(path)
	if not model.is_initialized
      puts "#{results_directory} has not been run"
      next
	end	
	model = model.get

  # get and shorten building name
  building_name = model.getBuilding.name.to_s
  puts "#{building_name} is in directory (#{results_directory})"
  dash_index = building_name.index('-')
	if not dash_index.nil?
		short_name = building_name[0,dash_index - 1]
  else
    short_name = building_name
  end

	# copy and rename zip file
	orig_zip = "#{File.dirname(__FILE__)}/#{results_directory}/run/data_point.zip"
	copy_zip = "../results/bestest_zips/#{short_name}.zip"
	puts "Creating #{copy_zip}"
	directory_name = "../results/bestest_zips"
	Dir.mkdir(directory_name) unless File.exists?(directory_name)
	FileUtils.cp(orig_zip, copy_zip)

end
