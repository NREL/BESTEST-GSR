# This takes data from OpenStudio server csv file and populates a copy of the Standard 140 Results spreadsheets
# See steps below taken prior to running this script.

# Run OpenStudio server projects from "integration testing directory"
# The reporting measure in the project contains runner.registerValues objects that in turn get written into the results csv.
# In the future the runner.registerValue data will live in the OSW file with each datapoint.
# run scrpint from directory script is in "Results"

# requires
require 'csv'
require 'fileutils'
require 'rubyXL' # install gem first
# gem documentation # http://www.rubydoc.info/gems/rubyXL/1.1.12/RubyXL/Cell
# https://github.com/weshatheleopard/rubyXL
require "#{File.dirname(__FILE__)}/resources/common_info"

# array for historical rows
historical_gen_info = []
historical_rows = []

# Load in CSV file
#csv_file = 'PAT_BESTEST_HE.csv'
csv_file = 'workflow_results.csv' # bestest.case_num will be first column trip for header

csv_hash = {}
CSV.foreach(csv_file, :headers => true, :header_converters => :symbol, :converters => :all) do |row|
  short_name = row.fields[0].split(" ").first
  csv_hash[short_name] = Hash[row.headers[1..-1].zip(row.fields[1..-1])]
end
puts "CSV has #{csv_hash.size} entries."
puts "Hash keys are #{csv_hash.keys}" # keys made from column 6

# Copy Excel File
orig_results_5_2a = 'resources/RESULTS5-2A.xlsx'
copy_results_5_2a = 'RESULTS5-2A.xlsx'
puts "Making a copy of #{orig_results_5_2a}"
FileUtils.cp(orig_results_5_2a, copy_results_5_2a)

# Load Excel File
workbook = RubyXL::Parser.parse(copy_results_5_2a)
worksheet = workbook['YourData']
puts "Loading #{worksheet.sheet_name} Worksheet"

category = "Annual Heating Loads"
puts "Populating #{category}"
(64..98).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingannual_heating])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
end

category = "Annual Cooling Loads"
puts "Populating #{category}"
(103..137).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingannual_cooling])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
end

category = "Annual Houlry Integrated Peak Heating Loads"
puts "Populating #{category}"
(145..179).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  puts "working on case #{target_case}"

  # get date and time from raw value
  raw_value = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_heating_time]
  date = raw_value[0,6]
  time = raw_value[7,2].to_i

  # populate value date and time columns
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_heating_value])
  worksheet.sheet_data[i][2].change_contents(date)
  worksheet.sheet_data[i][3].change_contents(time)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[144][1].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[144][2].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[144][3].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
end

category = "Annual Houlry Integrated Peak Cooling Loads"
puts "Populating #{category}"
(198..232).each do |i|
  target_case = worksheet.sheet_data[i][0].value

  # get date and time from raw value
  raw_value = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_cooling_time]
  date = raw_value[0,6]
  time = raw_value[7,2].to_i

  # populate value date and time columns
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_cooling_value])
  worksheet.sheet_data[i][2].change_contents(date)
  worksheet.sheet_data[i][3].change_contents(time)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[197][1].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[197][2].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[197][3].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
end

# date format should be dd-MMM. Hour is integer
# todo - would be nice to redo this to use process_output_timeseries in reporting measure to get time directly
def self.return_date_time_from_8760_index(index)

  date_string = nil
  dd = nil
  mmm = nil
  hour = nil

  # assuming non leap year
  month_hash = {}
  month_hash['JAN'] = 31
  month_hash['FEB'] = 28
  month_hash['MAR'] = 31
  month_hash['APR'] = 30
  month_hash['MAY'] = 31
  month_hash['JUN'] = 30
  month_hash['JUL'] = 31
  month_hash['AUG'] = 31
  month_hash['SEP'] = 30
  month_hash['OCT'] = 31
  month_hash['NOV'] = 30
  month_hash['DEC'] = 31

  raw_date = (index/24.0).floor
  counter = 0
  month_hash.each do |k,v|
    if raw_date - counter <= v
      # found month
      mmm = k
      dd = 1 + raw_date - counter
      date_string = "#{"%02d" % dd}-#{mmm}"
      hour = (index % 24)
      return [date_string,hour]
    else
      counter = counter + v
    end
  end
  return nil # shouldn't hit this
end

# tag date and time
category = "FF Max Hourly Zone Temperature"
puts "Populating #{category}"
# this also includes case 960
(253..257).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmax_temp])
  index_position = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmax_index_position]
  date_time_array = return_date_time_from_8760_index(index_position)
  worksheet.sheet_data[i][2].change_contents(date_time_array[0])
  worksheet.sheet_data[i][3].change_contents(date_time_array[1])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[252][1].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[252][2].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[252][3].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
end

# tag date and time
category = "FF Min Hourly Zone Temperature"
puts "Populating #{category}"
# this also includes case 960
(262..266).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  # populate value date and time columns
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmin_temp])
  index_position = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmin_index_position]
  date_time_array = return_date_time_from_8760_index(index_position)
  worksheet.sheet_data[i][2].change_contents(date_time_array[0])
  worksheet.sheet_data[i][3].change_contents(date_time_array[1])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[261][1].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[261][2].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}_#{worksheet.sheet_data[261][3].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
end

category = "FF Average Hourly Zone Temperature"
puts "Populating #{category}"
# this also includes case 960
(271..275).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  # populate value date and time columns
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingavg_temp])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
end

category = "Annual Incident Total Case 600"
puts "Populating #{category}"
target_case = '600'
worksheet.sheet_data[293][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingnorth_incident_solar_radiation])
worksheet.sheet_data[294][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingeast_incident_solar_radiation])
worksheet.sheet_data[295][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingwest_incident_solar_radiation])
worksheet.sheet_data[296][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingsouth_incident_solar_radiation])
worksheet.sheet_data[297][1].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportinghorizontal_incident_solar_radiation])
historical_rows << ["#{category} #{worksheet.sheet_data[293][0].value.to_s}",worksheet.sheet_data[293][1].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[294][0].value.to_s}",worksheet.sheet_data[294][1].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[295][0].value.to_s}",worksheet.sheet_data[295][1].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[296][0].value.to_s}",worksheet.sheet_data[296][1].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[297][0].value.to_s}",worksheet.sheet_data[297][1].value.to_s]

# changing cases not to match what
category = "Unshaded Annual Transmitted Cases 620 and 600"
puts "Populating #{category}"
worksheet.sheet_data[312][1].change_contents(csv_hash['620'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
worksheet.sheet_data[313][1].change_contents(csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
historical_rows << ["#{category} #{worksheet.sheet_data[312][0].value.to_s}",worksheet.sheet_data[312][1].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[313][0].value.to_s}",worksheet.sheet_data[313][1].value.to_s]

category = "Shaded Annual Transmitted Cases 930 and 910"
puts "Populating #{category}"
worksheet.sheet_data[332][1].change_contents(csv_hash['930'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
worksheet.sheet_data[333][1].change_contents(csv_hash['910'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
historical_rows << ["#{category} #{worksheet.sheet_data[332][0].value.to_s}",worksheet.sheet_data[332][1].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[333][0].value.to_s}",worksheet.sheet_data[333][1].value.to_s]

category = "Hourly Incident Solar Radiation Cloudy Day March 5th Case 600 - South"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0305_zone_surface_south].split(",")
counter = 0
(348..371).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Cloudy Day March 5th Case 600 - West"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0305_zone_surface_west].split(",")
counter = 0
(388..411).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Clear Day July 27th Case 600 - South"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0727_zone_surface_south].split(",")
counter = 0
(428..451).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Clear Dat July 27th Case 600 - West"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0727_zone_surface_west].split(",")
counter = 0
(468..491).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly FF Temperatures January 4th - Case 600FF"
puts "Populating #{category}"
array = csv_hash['600FF'][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0104].split(",")
counter = 0
(507..530).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+1].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly FF Temperatures January 4th - Case 900FF"
puts "Populating #{category}"
array = csv_hash['900FF'][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0104].split(",")
counter = 0
(547..570).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+1].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly FF Temperatures July 27 - Case 650FF"
puts "Populating #{category}"
array = csv_hash['650FF'][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0727].split(",")
counter = 0
(587..610).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+1].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly FF Temperatures July 27 - Case 950FF"
puts "Populating #{category}"
array = csv_hash['950FF'][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0727].split(",")
counter = 0
(627..650).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+1].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly Heating and Cooling Load 0104 - Case 600"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsens_htg_clg_0104].split(",")
counter = 0
(667..690).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly Heating and Cooling Load 0104 - Case 900"
puts "Populating #{category}"
array = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingsens_htg_clg_0104].split(",")
counter = 0
(707..730).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

category = "Hourly Annual Zone Temperature Bin Data - Case 900FF"
puts "Populating #{category}"
array = csv_hash['900FF'][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_bins].split(",")
# bin array is just -20 to 70C. The spreadsheet looks for -50 to 98C. May need to extend array or make blanks 0.
counter = 0
(779..868).each do |i|
  worksheet.sheet_data[i][1].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s}",worksheet.sheet_data[i][1].value.to_s]
  counter += 1
end

puts "Adding General Information"
# gather general information
common_info = BestestResults.populate_common_info

# starting position
gen_info_row = 45
gen_info_col = 0

# populate generalinfo
worksheet.sheet_data[gen_info_row][gen_info_col].change_contents(common_info[:program_name_and_version])
worksheet.sheet_data[gen_info_row+1][gen_info_col+4].change_contents(common_info[:program_version_release_date])
worksheet.sheet_data[gen_info_row+2][gen_info_col+4].change_contents(common_info[:program_name_short])
worksheet.sheet_data[gen_info_row+3][gen_info_col+4].change_contents(common_info[:results_submission_date])
# row skiped in Excel
worksheet.sheet_data[gen_info_row+5][gen_info_col].change_contents(common_info[:organization])
worksheet.sheet_data[gen_info_row+6][gen_info_col+4].change_contents(common_info[:organization_short])

# add general info to historical file
historical_gen_info << ["program_name_and_version",common_info[:program_name_and_version]]
historical_gen_info << ["program_version_release_date",common_info[:program_version_release_date]]
historical_gen_info << ["program_name_short",common_info[:program_name_short]]
historical_gen_info << ["results_submission_date",common_info[:results_submission_date]]
historical_gen_info << ["organization",common_info[:organization]]
historical_gen_info << ["organization_short",common_info[:organization_short]]

# Save Updated Excel File
puts "Saving #{copy_results_5_2a}"
workbook.write(copy_results_5_2a)

# create OpenStudio copy with updated program info
# Copy Excel File
os_copy_results_5_2a = 'RESULTS5-2a_OS.xlsx'
puts "Making an OpenStudio copy of #{copy_results_5_2a}"
FileUtils.cp(copy_results_5_2a, os_copy_results_5_2a)

puts "Adding General Information"
# gather general information
common_info = BestestResults.populate_common_info("OS")

# starting position
gen_info_row = 45
gen_info_col = 0

# populate generalinfo
worksheet.sheet_data[gen_info_row][gen_info_col].change_contents(common_info[:program_name_and_version])
worksheet.sheet_data[gen_info_row+1][gen_info_col+4].change_contents(common_info[:program_version_release_date])
worksheet.sheet_data[gen_info_row+2][gen_info_col+4].change_contents(common_info[:program_name_short])
worksheet.sheet_data[gen_info_row+3][gen_info_col+4].change_contents(common_info[:results_submission_date])
# row skiped in Excel
worksheet.sheet_data[gen_info_row+5][gen_info_col].change_contents(common_info[:organization])
worksheet.sheet_data[gen_info_row+6][gen_info_col+4].change_contents(common_info[:organization_short])

# Save Updated Excel File
puts "Saving #{os_copy_results_5_2a}"
workbook.write(os_copy_results_5_2a)

# load CSV file with historical version results
historical_file = "historical/#{common_info[:program_name_and_version].gsub(".","_").gsub(" ","_")}.csv"
puts "Saving #{historical_file}"
CSV.open(historical_file, "w") do |csv|
  [*historical_gen_info,*historical_rows].each do |row|
    csv << row
  end
end