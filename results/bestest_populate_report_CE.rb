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
require 'rubyXL/convenience_methods'
# gem documentation # http://www.rubydoc.info/gems/rubyXL/1.1.12/RubyXL/Cell
# https://github.com/weshatheleopard/rubyXL
require "#{File.dirname(__FILE__)}/resources/common_info"

# array for historical rows
historical_gen_info = []
historical_rows = []

# Load in CSV file
csv_file = 'workflow_results.csv' # bestest.case_num will be first column trip for header
csv_hash = {}
CSV.foreach(csv_file, :headers => true, :header_converters => :symbol, :converters => :all) do |row|
  short_name = row.fields[0].split(" ").first
  csv_hash[short_name] = Hash[row.headers[1..-1].zip(row.fields[1..-1])]
end
puts "CSV has #{csv_hash.size} entries."
puts "Hash keys are #{csv_hash.keys}" # keys made from column 6

# todo - before I commit update all registerValues ot be unique (add table name and if needed 3a vs. 3b)

# Copy first Excel File
orig_results_5_3a = 'resources/RESULTS5-3A.xlsx'
copy_results_5_3a = 'RESULTS5-3A.xlsx'
puts "Making a copy of #{orig_results_5_3a}"
FileUtils.cp(orig_results_5_3a, copy_results_5_3a)

# Load Excel File
workbook = RubyXL::Parser.parse(copy_results_5_3a)
worksheet = workbook['YourData']
puts "Loading #{worksheet.sheet_name} Worksheet"

# make array for columns on table
columns = []
columns << :bestest_ce_reportingclg_energy_consumption_total
columns << :bestest_ce_reportingclg_energy_consumption_compressor
columns << :bestest_ce_reportingclg_energy_consumption_supply_fan
columns << :bestest_ce_reportingclg_energy_consumption_condenser_fan
columns << :bestest_ce_reportingevaporator_coil_load_total
columns << :bestest_ce_reportingevaporator_coil_load_sensible
columns << :bestest_ce_reportingevaporator_coil_load_latent
columns << :bestest_ce_reportingzone_load_total
columns << :bestest_ce_reportingzone_load_sensible
columns << :bestest_ce_reportingzone_load_latent
columns << :bestest_ce_reportingfeb_mean_cop
columns << :bestest_ce_reportingfeb_mean_idb
columns << :bestest_ce_reportingfeb_mean_humidity_ratio
columns << :bestest_ce_reportingfeb_max_cop
columns << :bestest_ce_reportingfeb_max_idb
columns << :bestest_ce_reportingfeb_max_humidity_ratio
columns << :bestest_ce_reportingfeb_min_cop
columns << :bestest_ce_reportingfeb_min_idb
columns << :bestest_ce_reportingfeb_min_humidity_ratio

# populate table on YourData
puts "Populating main table for 5-3A"
(24..37).each do |i|
  target_case = worksheet.sheet_data[i][0].value

  puts "Adding row for #{target_case}"
  # loop through columns for each case
  columns.each_with_index do |column,j|

    # skip specifc columns that can't be calculated
    next if column == :bestest_ce_reportingclg_energy_consumption_compressor
    next if column == :bestest_ce_reportingclg_energy_consumption_condenser_fan

    worksheet.sheet_data[i][j+1].change_contents(csv_hash[target_case][column])
    historical_rows << ["#{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+1].value.to_s]
  end

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

# add general info to historical file
historical_gen_info << ["program_name_and_version",common_info[:program_name_and_version]]
historical_gen_info << ["program_version_release_date",common_info[:program_version_release_date]]
historical_gen_info << ["program_name_short",common_info[:program_name_short]]
historical_gen_info << ["results_submission_date",common_info[:results_submission_date]]
historical_gen_info << ["organization",common_info[:organization]]
historical_gen_info << ["organization_short",common_info[:organization_short]]

# Save Updated Excel File
# todo - Sheet A that refers to YourData isn't updating cells that refer to YourData. Not sure why. Workaround for now is to copy and paste YourData when I first open it, but shouldn't have to do that.
puts "Saving #{copy_results_5_3a}"
workbook.write(copy_results_5_3a)


# create OpenStudio copy with updated program info
# Copy Excel File
os_copy_results_5_3a = 'RESULTS5-3A_OS.xlsx'
puts "Making an OpenStudio copy of #{copy_results_5_3a}"
FileUtils.cp(copy_results_5_3a, os_copy_results_5_3a)

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
puts "Saving #{os_copy_results_5_3a}"
workbook.write(os_copy_results_5_3a)


# Copy second Excel File
orig_results_5_3b = 'resources/RESULTS5-3B.xlsx'
copy_results_5_3b = 'RESULTS5-3B.xlsx'
puts "Making a copy of #{orig_results_5_3b}"
FileUtils.cp(orig_results_5_3b, copy_results_5_3b)

# Load Excel File
workbook = RubyXL::Parser.parse(copy_results_5_3b)
worksheet = workbook['YourData']
puts "Loading #{worksheet.sheet_name} Worksheet"

# make array for columns on table
columns = []

columns << :bestest_ce_reportingann_sum_clg_energy_consumption_total
columns << :bestest_ce_reportingann_sum_clg_energy_consumption_compressor
columns << :bestest_ce_reportingann_sum_clg_energy_consumption_condenser_fan
columns << :bestest_ce_reportingann_sum_clg_energy_consumption_supply_fan
columns << :bestest_ce_reportingann_sum_evap_coil_load_total
columns << :bestest_ce_reportingann_sum_evap_coil_load_sensible
columns << :bestest_ce_reportingann_sum_evap_coil_load_latent
columns << :bestest_ce_reportingann_mean_cop_2
columns << :bestest_ce_reportingann_mean_idb
columns << :bestest_ce_reportingann_mean_zone_humidity_ratio
columns << :bestest_ce_reportingann_mean_zone_relative_humidity

columns_extra_300 = []
columns_extra_300 << :bestest_ce_reportingann_mean_odb # CE300 only
columns_extra_300 << :bestest_ce_reportingann_mean_outdoor_humidity_ratio #CE300 only

# populate table on YourData
puts "Populating Annual Sums and Means Table"
(61..73).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  puts "Adding row for #{target_case}"
  # loop through columns for each case
  columns.each_with_index do |column,j|

    # skip specifc columns that can't be calculated
    next if column == :bestest_ce_reportingann_sum_clg_energy_consumption_compressor
    next if column == :bestest_ce_reportingann_sum_clg_energy_consumption_condenser_fan

    worksheet.sheet_data[i][j+1].change_contents(csv_hash[target_case][column])
    historical_rows << ["#{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+1].value.to_s]
  end
  # extra columns just for CE300
  if target_case.include? 'CE300'
    columns_extra_300.each_with_index do |column,j|

      worksheet.sheet_data[i][j+12].change_contents(csv_hash[target_case][column])
      historical_rows << ["#{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+12].value.to_s]
    end
  end
end
(76..81).each do |i|
  target_case = worksheet.sheet_data[i][0].value
  puts "Adding row for #{target_case}"
  # loop through columns for each case
  columns.each_with_index do |column,j|

    # skip specifc columns that can't be calculated
    next if column == :bestest_ce_reportingann_sum_clg_energy_consumption_compressor
    next if column == :bestest_ce_reportingann_sum_clg_energy_consumption_condenser_fan

    worksheet.sheet_data[i][j+1].change_contents(csv_hash[target_case][column])
    historical_rows << ["#{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+1].value.to_s]
  end
end

# Fill in two special rows in table for CE500 and CE510 only
columns = []
columns << :bestest_ce_reportingmay_sept_sum_clg_consumption_total
columns << :bestest_ce_reportingmay_sept_sum_clg_consumption_compressor
columns << :bestest_ce_reportingmay_sept_sum_clg_consumption_cond_fan
columns << :bestest_ce_reportingmay_sept_sum_clg_consumption_indoor_fan
columns << :bestest_ce_reportingmay_sept_sum_evap_coil_load_total
columns << :bestest_ce_reportingmay_sept_sum_evap_coil_load_sensible
columns << :bestest_ce_reportingmay_sept_sum_evap_coil_load_latent
columns << :bestest_ce_reportingmay_sept_mean_cop_2
columns << :bestest_ce_reportingmay_sept_mean_idb
columns << :bestest_ce_reportingmay_sept_mean_zone_humidity_ratio
columns << :bestest_ce_reportingmay_sept_mean_zone_relative_humidity

# CE500 May-Sep
columns.each_with_index do |column,j|

  # skip specifc columns that can't be calculated
  next if column == :bestest_ce_reportingmay_sept_sum_clg_consumption_compressor
  next if column == :bestest_ce_reportingmay_sept_sum_clg_consumption_cond_fan

  worksheet.sheet_data[74][j+1].change_contents(csv_hash["CE500"][column])
  historical_rows << ["#{worksheet.sheet_data[74][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[74][j+1].value.to_s]
end
# CE510 May-Sep
columns.each_with_index do |column,j|

  # skip specifc columns that can't be calculated
  next if column == :bestest_ce_reportingmay_sept_sum_clg_consumption_compressor
  next if column == :bestest_ce_reportingmay_sept_sum_clg_consumption_cond_fan

  worksheet.sheet_data[75][j+1].change_contents(csv_hash["CE510"][column])
  historical_rows << ["#{worksheet.sheet_data[75][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[75][j+1].value.to_s]
end

# make array for columns on table
columns = []
columns << :bestest_ce_reportingenergy_consumption_comp_both_fans_wh
columns << :bestest_ce_reportingenergy_consumption_comp_both_fans_date
columns << :bestest_ce_reportingenergy_consumption_comp_both_fans_hr
columns << :bestest_ce_reportingevap_coil_load_sensible_wh
columns << :bestest_ce_reportingevap_coil_load_sensible_date
columns << :bestest_ce_reportingevap_coil_load_sensible_hr
columns << :bestest_ce_reportingevap_coil_load_latent_wh
columns << :bestest_ce_reportingevap_coil_load_latent_date
columns << :bestest_ce_reportingevap_coil_load_latent_hr
columns << :bestest_ce_reportingevap_coil_load_sensible_and_latent_wh
columns << :bestest_ce_reportingevap_coil_load_sensible_and_latent_date
columns << :bestest_ce_reportingevap_coil_load_sensible_and_latent_hr

columns_extra_300 = []
columns_extra_300 << :bestest_ce_reportingweather_odb_c # CE300 only
columns_extra_300 << :bestest_ce_reportingweather_odb_date # CE300 only
columns_extra_300 << :bestest_ce_reportingweather_odb_hr # CE300 only
columns_extra_300 << :bestest_ce_reportingweather_outdoor_humidity_ratio_c # CE300 only
columns_extra_300 << :bestest_ce_reportingweather_outdoor_humidity_ratio_date # CE300 only
columns_extra_300 << :bestest_ce_reportingweather_outdoor_humidity_ratio_hr # CE300 only

# populate table on YourData
puts "Populating Annual Hourly Integrated Maxima Consumptions and Loads Table"
(61..80).each do |i|
  target_case = worksheet.sheet_data[i][15].value
  if target_case.include? "CE500" then target_case = "CE500" end # raw spreadsheet has extra space in cell
  # loop through columns for each case
  columns.each_with_index do |column,j|
    worksheet.sheet_data[i][j+16].change_contents(csv_hash[target_case][column])
    historical_rows << ["#{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+16].value.to_s]
  end
  # extra columns just for CE300
  if target_case.include? 'CE300'
    columns_extra_300.each_with_index do |column,j|
      worksheet.sheet_data[i][j+28].change_contents(csv_hash[target_case][column])
      historical_rows << ["CE300 #{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+28].value.to_s]
    end
  end
end

# make array for columns on table
columns = []
columns << :bestest_ce_reportingcop_2_max_cop_2
columns << :bestest_ce_reportingcop_2_max_date
columns << :bestest_ce_reportingcop_2_max_hr
columns << :bestest_ce_reportingcop_2_min_cop_2
columns << :bestest_ce_reportingcop_2_min_date
columns << :bestest_ce_reportingcop_2_min_hr
columns << :bestest_ce_reportingidb_max_idb
columns << :bestest_ce_reportingidb_max_date
columns << :bestest_ce_reportingidb_max_hr
columns << :bestest_ce_reportingidb_min_idb
columns << :bestest_ce_reportingidb_min_date
columns << :bestest_ce_reportingidb_min_hr
columns << :bestest_ce_reportinghr_max_humidity_ratio
columns << :bestest_ce_reportinghr_max_date
columns << :bestest_ce_reportinghr_max_hr
columns << :bestest_ce_reportinghr_min_humidity_ratio
columns << :bestest_ce_reportinghr_min_date
columns << :bestest_ce_reportinghr_min_hr
columns << :bestest_ce_reportingrh_max_relative_humidity
columns << :bestest_ce_reportingrh_max_date
columns << :bestest_ce_reportingrh_max_hr
columns << :bestest_ce_reportingrh_min_relative_humidity
columns << :bestest_ce_reportingrh_min_date
columns << :bestest_ce_reportingrh_min_hr

# populate table on YourData
category = "Annual Hourly Integrated Maxima - cop_2 and Zone Table"
puts "Populating #{category}"
(88..99).each do |i|
  target_case = worksheet.sheet_data[i][15].value
  puts "Adding row for #{target_case}"
  # loop through columns for each case
  columns.each_with_index do |column,j|
    worksheet.sheet_data[i][j+16].change_contents(csv_hash[target_case][column])
    historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+16].value.to_s]
  end
end

# make array for columns on table
columns = []
columns << :bestest_ce_reportingapr_dec_cop_2_max_cop_2
columns << :bestest_ce_reportingapr_dec_cop_2_max_date
columns << :bestest_ce_reportingapr_dec_cop_2_max_hr
columns << :bestest_ce_reportingapr_dec_cop_2_min_cop_2
columns << :bestest_ce_reportingapr_dec_cop_2_min_date
columns << :bestest_ce_reportingapr_dec_cop_2_min_hr
columns << :bestest_ce_reportingapr_dec_idb_max_idb
columns << :bestest_ce_reportingapr_dec_idb_max_date
columns << :bestest_ce_reportingapr_dec_idb_max_hr
columns << :bestest_ce_reportingapr_dec_idb_min_idb
columns << :bestest_ce_reportingapr_dec_idb_min_date
columns << :bestest_ce_reportingapr_dec_idb_min_hr
columns << :bestest_ce_reportingapr_dec_hr_max_humidity_ratio
columns << :bestest_ce_reportingapr_dec_hr_max_date
columns << :bestest_ce_reportingapr_dec_hr_max_hr
columns << :bestest_ce_reportingapr_dec_hr_min_humidity_ratio
columns << :bestest_ce_reportingapr_dec_hr_min_date
columns << :bestest_ce_reportingapr_dec_hr_min_hr
columns << :bestest_ce_reportingapr_dec_rh_max_relative_humidity
columns << :bestest_ce_reportingapr_dec_rh_max_date
columns << :bestest_ce_reportingapr_dec_rh_max_hr
columns << :bestest_ce_reportingapr_dec_rh_min_relative_humidity
columns << :bestest_ce_reportingapr_dec_rh_min_date
columns << :bestest_ce_reportingapr_dec_rh_min_hr

# populate table on YourData
category =  "Annual Hourly Integrated Maxima - cop_2 and Zone Table"
puts "Populating #{category}"
(100..107).each do |i|
  target_case = worksheet.sheet_data[i][15].value
  if target_case.include? "CE500" then target_case = "CE500" end # raw spreadsheet has extra space in cell
  puts "Adding row for #{target_case}"
  # loop through columns for each case
  columns.each_with_index do |column,j|
    worksheet.sheet_data[i][j+16].change_contents(csv_hash[target_case][column])
    historical_rows << ["#{category} #{worksheet.sheet_data[i][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i][j+16].value.to_s]
  end
end

# pouplate table
# each column is registerValue with string that can be converted to array with 24 items
columns = []
columns << :bestest_ce_reportingmmdd_0628_hourly_energy_consumpton_compressor
columns << :bestest_ce_reportingmmdd_0628_hourly_energy_consumpton_cond_fan
columns << :bestest_ce_reportingmmdd_0628_hourly_evaporator_coil_load_total
columns << :bestest_ce_reportingmmdd_0628_hourly_evaporator_coil_load_sensible
columns << :bestest_ce_reportingmmdd_0628_hourly_evaporator_coil_load_latent
columns << :bestest_ce_reportingmmdd_0628_hourly_zone_humidity_ratio
columns << :bestest_ce_reportingmmdd_0628_hourly_cop_2
columns << :bestest_ce_reportingmmdd_0628_hourly_odb
columns << :bestest_ce_reportingmmdd_0628_hourly_edb
columns << :bestest_ce_reportingmmdd_0628_hourly_ewb
columns << :bestest_ce_reportingmmdd_0628_hourly_outdoor_humidity_ratio

category =  "Case 300 June 28th Hourly Table"
puts "Populating #{category}"
# todo - convert string to number for each cell
columns.each_with_index do |column,j|
  array = csv_hash['CE300'][column].split(',')
  array.each_with_index do |hourly_value,i|

    # skip specifc columns that can't be calculated
    next if column == :bestest_ce_reportingmmdd_0628_hourly_energy_consumpton_cond_fan

    if not hourly_value == "tbd"
      hourly_value = hourly_value.to_f
    end
    worksheet.sheet_data[i+88][j+1].change_contents(hourly_value)
    historical_rows << ["#{category} #{worksheet.sheet_data[i+88][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[i+88][j+1].value.to_s]
  end
end

# make array for columns on table
columns = []
columns << :bestest_ce_reportingmmdd_0430_day_energy_consumption_total
columns << :bestest_ce_reportingmmdd_0430_day_energy_consumption_compressor
columns << :bestest_ce_reportingmmdd_0430_day_energy_consumption_condenser_fan
columns << :bestest_ce_reportingmmdd_0430_day_energy_consumption_supply_fan
columns << :bestest_ce_reportingmmdd_0430_day_evaporator_coil_load_total
columns << :bestest_ce_reportingmmdd_0430_day_evaporator_coil_load_sensible
columns << :bestest_ce_reportingmmdd_0430_day_evaporator_coil_load_latent
columns << :bestest_ce_reportingmmdd_0430_day_zone_humidity_ratio
columns << :bestest_ce_reportingmmdd_0430_day_cop_2
columns << :bestest_ce_reportingmmdd_0430_day_odb
columns << :bestest_ce_reportingmmdd_0430_day_edb

# pouplate table
category =  "Case 500 and 530 Average Daily Outputs"
puts "Populating #{category}"
columns.each_with_index do |column,j|

  # skip specifc columns that can't be calculated
  next if column == :bestest_ce_reportingmmdd_0430_day_energy_consumption_compressor
  next if column == :bestest_ce_reportingmmdd_0430_day_energy_consumption_condenser_fan

  worksheet.sheet_data[119][j+1].change_contents(csv_hash['CE500'][column])
  worksheet.sheet_data[128][j+1].change_contents(csv_hash['CE530'][column])
  historical_rows << ["#{category} #{worksheet.sheet_data[119][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[119][j+1].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[128][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[128][j+1].value.to_s]

end

# make array for columns on table
columns = []
columns << :bestest_ce_reportingmmdd_0625_day_energy_consumption_total
columns << :bestest_ce_reportingmmdd_0625_day_energy_consumption_compressor
columns << :bestest_ce_reportingmmdd_0625_day_energy_consumption_condenser_fan
columns << :bestest_ce_reportingmmdd_0625_day_energy_consumption_supply_fan
columns << :bestest_ce_reportingmmdd_0625_day_evaporator_coil_load_total
columns << :bestest_ce_reportingmmdd_0625_day_evaporator_coil_load_sensible
columns << :bestest_ce_reportingmmdd_0625_day_evaporator_coil_load_latent
columns << :bestest_ce_reportingmmdd_0625_day_zone_humidity_ratio
columns << :bestest_ce_reportingmmdd_0625_day_cop_2
columns << :bestest_ce_reportingmmdd_0625_day_odb
columns << :bestest_ce_reportingmmdd_0625_day_edb

columns.each_with_index do |column,j|

  # skip specifc columns that can't be calculated
  next if column == :bestest_ce_reportingmmdd_0625_day_energy_consumption_compressor
  next if column == :bestest_ce_reportingmmdd_0625_day_energy_consumption_condenser_fan

  worksheet.sheet_data[120][j+1].change_contents(csv_hash['CE500'][column])
  worksheet.sheet_data[129][j+1].change_contents(csv_hash['CE530'][column])
  historical_rows << ["#{category} #{worksheet.sheet_data[120][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[120][j+1].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[129][0].value.to_s} #{column.to_s.gsub("bestest_ce_reporting","")}",worksheet.sheet_data[129][j+1].value.to_s]
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
puts "Saving #{copy_results_5_3b}"
workbook.write(copy_results_5_3b)

# create OpenStudio copy with updated program info
# Copy Excel File
os_copy_results_5_3b = 'RESULTS5-3B_OS.xlsx'
puts "Making an OpenStudio copy of #{copy_results_5_3b}"
FileUtils.cp(copy_results_5_3b, os_copy_results_5_3b)

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
puts "Saving #{os_copy_results_5_3b}"
workbook.write(os_copy_results_5_3b)

# load CSV file with historical version results
historical_file = "historical/#{common_info[:program_name_and_version].gsub(".","_").gsub(" ","_")}_CE.csv"
puts "Saving #{historical_file}"
CSV.open(historical_file, "w") do |csv|
  [*historical_gen_info,*historical_rows].each do |row|
    csv << row
  end
end