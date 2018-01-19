require 'fileutils'
require 'openstudio'

# expects single argument with path to directory that contains dataoints
#path_datapoints = 'run/MyProject'
#path_datapoints = ARGV[0]
path_datapoints = "workflow"

# todo - create a hash to contain all results data vs. simple CSV rows
results_hash = {}

# loop through resoruce files
results_directories = Dir.glob("#{path_datapoints}/*")
results_directories.each do |results_directory|

  next if not results_directory.include?("BESTEST Case ")

  row_data = {}

	# load the test model
  # todo - update to get from idf vs. osm
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

  dash_index = building_name.index('-')
	if ! dash_index.nil?
	  short_name = building_name[0,dash_index - 1]
  else
    short_name = building_name
  end
  puts "#{short_name} is in directory (#{results_directory})"

  # load OSW to get information from argument values
  osw_path = OpenStudio::Path.new("#{File.dirname(__FILE__)}/#{results_directory}/out.osw")
  osw = OpenStudio::WorkflowJSON.load(osw_path).get
  runner = OpenStudio::Measure::OSRunner.new(osw)

  # store high level information about datapoint
  row_data["dir_name"] = results_directory
  row_data["building_name"] = short_name
  #row_data["description"] = runner.workflow.description
  row_data["status"] = runner.workflow.completedStatus.get

  runner.workflow.workflowSteps.each do |step|
    if step.to_MeasureStep.is_initialized

      measure_step = step.to_MeasureStep.get
      measure_dir_name = measure_step.measureDirName

      if measure_step.name.is_initialized

        # for manual PAT projects I want to pass in the measure dir name as the header instead of the measure opiton name
        measure_step_name = measure_dir_name.downcase.gsub(" ","_").to_sym
        # measure_step_name = measure_step.name.get.downcase.gsub(" ","_").to_sym

        next if ! measure_step.result.is_initialized
        next if ! measure_step.result.get.stepResult.is_initialized
        measure_step_result = measure_step.result.get.stepResult.get.valueName

        # populate registerValue objects
        result = measure_step.result.get
        next if result.stepValues.size == 0
        #row_data[measure_step_name] = measure_step_result
        result.stepValues.each do |value|
          # populate feature_hash (there is issue filed with value.units)
          if value.name == "case_num"
            row_data["bestest.#{value.name}"] = value.valueAsVariant.to_s
          else
            row_data["#{measure_step_name}.#{value.name}"] = value.valueAsVariant.to_s
          end
        end

        # populate results_hash
        results_hash[results_directory] = row_data

      end

    else
      #puts "This step is not a measure"
    end

  end

end

# populate csv header
headers = []
results_hash.each do |k,v|
  v.each do |k2,v2|
    if ! headers.include? k2
      headers << k2
    end
  end
end
headers = headers.sort

# populate csv
require "csv"
csv_rows = []
results_hash.each do |k,v|
  arr_row = []
  headers.each {|header| arr_row.push(v.key?(header) ? v[header] : nil)}
  csv_row = CSV::Row.new(headers, arr_row)
  csv_rows.push(csv_row)  
end

# save csv
csv_table = CSV::Table.new(csv_rows)
path_report = "../results/workflow_results.csv"
puts "saving csv file to #{path_report}"
File.open(path_report, 'w'){|file| file << csv_table.to_s}