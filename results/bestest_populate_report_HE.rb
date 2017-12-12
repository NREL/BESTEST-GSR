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


# Load in CSV file
csv_file = 'PAT_BESTEST_HE.csv'
csv_hash = {}
CSV.foreach(csv_file, :headers => true, :header_converters => :symbol, :converters => :all) do |row|
  short_name = row.fields[6].split(" ").first
  csv_hash[short_name] = Hash[row.headers[1..-1].zip(row.fields[1..-1])]
end
puts "CSV has #{csv_hash.size} entries."
puts "Hash keys are #{csv_hash.keys}" # keys made from column 6

# Copy Excel File
orig_results_5_4 = 'resources/RESULTS5-4.xlsx'
copy_results_5_4 = 'RESULTS5-4.xlsx'
puts "Making a copy of #{orig_results_5_4}"
FileUtils.cp(orig_results_5_4, copy_results_5_4)

# Load Excel File
workbook = RubyXL::Parser.parse(copy_results_5_4)
worksheet = workbook['YourData']
puts "Loading #{worksheet.sheet_name} Worksheet"

puts "Populating Total Furnace Load"
(19..29).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingtotal_furnace_load])
end

puts "Populating Total Furnace Input"
(35..45).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingtotal_furnace_input])
end

# todo - change units here or in reporting measure
puts "Populating Fuel Consumption"
(51..61).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingaverage_fuel_consumption])
end

puts "Populating Fan Energy"
(67..72).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingfan_energy])
end

puts "Populating Mean Zone Temperature"
(78..80).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingmean_zone_temperature])
end

puts "Populating Maximum Zone Temperature"
(86..88).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingmaximum_zone_temperature])
end

puts "Populating Minimum Zone Temperature"
(94..96).each do |i|
  target_case = worksheet.sheet_data[i][0].value.to_s.split(':').first
  worksheet.sheet_data[i][1].change_contents(csv_hash[target_case][:bestest_he_reportingminimum_zone_temperature])
end

puts "Adding General Information"
# gather general information
common_info = BestestResults.populate_common_info

# starting position
gen_info_row = 1
gen_info_col = 5

# populate generalinfo
worksheet.sheet_data[gen_info_row][gen_info_col].change_contents(common_info[:program_name_and_version])
worksheet.sheet_data[gen_info_row+1][gen_info_col+4].change_contents(common_info[:program_version_release_date])
worksheet.sheet_data[gen_info_row+2][gen_info_col+4].change_contents(common_info[:program_name_short])
worksheet.sheet_data[gen_info_row+3][gen_info_col+4].change_contents(common_info[:results_submission_date])
# row skiped in Excel
worksheet.sheet_data[gen_info_row+5][gen_info_col].change_contents(common_info[:organization])
worksheet.sheet_data[gen_info_row+6][gen_info_col+4].change_contents(common_info[:organization_short])

# Save Updated Excel File
puts "Saving #{copy_results_5_4}"
workbook.write(copy_results_5_4)

# create OpenStudio copy with updated program info
# Copy Excel File
os_copy_results_5_4 = 'RESULTS5-4_OS.xlsx'
puts "Making an OpenStudio copy of #{copy_results_5_4}"
FileUtils.cp(copy_results_5_4, os_copy_results_5_4)

puts "Adding General Information"
# gather general information
common_info = BestestResults.populate_common_info("OS")

# starting position
gen_info_row = 1
gen_info_col = 5

# populate generalinfo
worksheet.sheet_data[gen_info_row][gen_info_col].change_contents(common_info[:program_name_and_version])
worksheet.sheet_data[gen_info_row+1][gen_info_col+4].change_contents(common_info[:program_version_release_date])
worksheet.sheet_data[gen_info_row+2][gen_info_col+4].change_contents(common_info[:program_name_short])
worksheet.sheet_data[gen_info_row+3][gen_info_col+4].change_contents(common_info[:results_submission_date])
# row skiped in Excel
worksheet.sheet_data[gen_info_row+5][gen_info_col].change_contents(common_info[:organization])
worksheet.sheet_data[gen_info_row+6][gen_info_col+4].change_contents(common_info[:organization_short])

# Save Updated Excel File
puts "Saving #{os_copy_results_5_4}"
workbook.write(os_copy_results_5_4)