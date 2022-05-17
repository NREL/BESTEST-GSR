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
(69..114).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s # new 2020 test cases for some reason are number while others are string, I had to add this.
  worksheet.sheet_data[i][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingannual_heating])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
end

category = "Annual Cooling Loads"
puts "Populating #{category}"
(69..114).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s
  worksheet.sheet_data[i][3].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingannual_cooling])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
end

category = "Annual Hourly Integrated Peak Heating Loads"
puts "Populating #{category}"
(69..114).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s

  # get date and time from raw value
  raw_value = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_heating_time]
  date = raw_value[0,6] 
  # takes month and day of month instead of date now)
  month = raw_value[3,3] # TODO - TYP confirm if month has to CamCase Capital (is JAN ok instead of Jan)
  day_of_month = raw_value[0,2] 
  # TODO - TYP investigate throughout this section, looking at std 140 2020 docs may want to increase hour to next hour (if time is x:mm where mm is anything but 00 then increase x by 1)
  time = raw_value[7,2].to_i

  # populate value date and time columns
  worksheet.sheet_data[i][4].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_heating_value])
  worksheet.sheet_data[i][5].change_contents(month)
  worksheet.sheet_data[i][6].change_contents(day_of_month)
  worksheet.sheet_data[i][7].change_contents(time)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[68][4].value.to_s}",worksheet.sheet_data[i][4].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_DATE}",date]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[68][7].value.to_s}",worksheet.sheet_data[i][7].value.to_s]
end

category = "Annual Hourly Integrated Peak Cooling Loads"
puts "Populating #{category}"
(69..114).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s

  # get date and time from raw value
  raw_value = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_cooling_time]
  date = raw_value[0,6]
  month = raw_value[3,3]
  day_of_month = raw_value[0,2]  
  time = raw_value[7,2].to_i

  # populate value date and time columns
  worksheet.sheet_data[i][8].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingpeak_cooling_value])
  worksheet.sheet_data[i][9].change_contents(date)
  worksheet.sheet_data[i][10].change_contents(time)
  worksheet.sheet_data[i][11].change_contents(time)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[68][8].value.to_s}",worksheet.sheet_data[i][8].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_DATE",date]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[68][11].value.to_s}",worksheet.sheet_data[i][11].value.to_s]
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

  raw_date = (index/24.0).ceil # updated from floor since first hour is 1 and nt 0
  counter = 0
  month_hash.each do |k,v|
    if raw_date - counter <= v
      # found month
      mmm = k
      dd = 1 + raw_date - counter
      date_string = "#{"%02d" % dd}-#{mmm}"
      hour = (index % 24)
      return [date_string,hour,mmm,dd] # kept date_sring for now and just added mmm and dd so I can use either for historical string for now.
    else
      counter = counter + v
    end
  end
  return nil # shouldn't hit this
end

# example string 28-NOV-12:00
# aything in first hour of day is 1 not 0
def self.return_date_time_from_string(string)

  date = {}
  date['dd'] = string[0..1].to_i
  date['mmm'] = string[3..5].to_i
  date['hh'] = string[7..8].to_i + 1

  return date

end

category = "FF Average Hourly Zone Temperature"
puts "Populating #{category}"
# this also includes case 960
(129..135).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s
  # populate value date and time columns
  worksheet.sheet_data[i][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingavg_temp])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
end

# tag date and time
category = "FF Min Hourly Zone Temperature"
puts "Populating #{category}"
# this also includes case 960
(129..135).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s
  # populate value date and time columns
  worksheet.sheet_data[i][3].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmin_temp])
  index_position = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmin_index_position]
  date_time_array = return_date_time_from_8760_index(index_position)
  worksheet.sheet_data[i][4].change_contents(date_time_array[2])
  worksheet.sheet_data[i][5].change_contents(date_time_array[3])
  worksheet.sheet_data[i][6].change_contents(date_time_array[1])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[128][3].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_DATE",date_time_array[0]]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[128][6].value.to_s}",worksheet.sheet_data[i][6].value.to_s]
end

# tag date and time
category = "FF Max Hourly Zone Temperature"
puts "Populating #{category}"
# this also includes case 960
(129..135).each do |i|
  target_case = worksheet.sheet_data[i][1].value.to_s
  worksheet.sheet_data[i][7].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmax_temp])
  index_position = csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingmax_index_position]
  date_time_array = return_date_time_from_8760_index(index_position)
  worksheet.sheet_data[i][8].change_contents(date_time_array[2])
  worksheet.sheet_data[i][9].change_contents(date_time_array[3])
  worksheet.sheet_data[i][10].change_contents(date_time_array[1])
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[128][7].value.to_s}",worksheet.sheet_data[i][7].value.to_s]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_DATE",date_time_array[0]]
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}_#{worksheet.sheet_data[128][10].value.to_s}",worksheet.sheet_data[i][10].value.to_s]
end

category = "Annual Incident Total Case 600"
puts "Populating #{category}"
target_case = '600'
# order of code is based on std 140 2017 but updated in place for std 140 2020
worksheet.sheet_data[155][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingnorth_incident_solar_radiation])
worksheet.sheet_data[156][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingeast_incident_solar_radiation])
worksheet.sheet_data[158][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingwest_incident_solar_radiation])
worksheet.sheet_data[157][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportingsouth_incident_solar_radiation])
worksheet.sheet_data[154][2].change_contents(csv_hash[target_case][:bestest_building_thermal_envelope_and_fabric_load_reportinghorizontal_incident_solar_radiation])
historical_rows << ["#{category} 600#{worksheet.sheet_data[155][1].value.to_s}",worksheet.sheet_data[155][2].value.to_s]
historical_rows << ["#{category} 600#{worksheet.sheet_data[156][1].value.to_s}",worksheet.sheet_data[156][2].value.to_s]
historical_rows << ["#{category} 600#{worksheet.sheet_data[158][1].value.to_s}",worksheet.sheet_data[158][2].value.to_s]
historical_rows << ["#{category} 600#{worksheet.sheet_data[157][1].value.to_s}",worksheet.sheet_data[157][2].value.to_s]
historical_rows << ["#{category} 600#{worksheet.sheet_data[154][1].value.to_s}",worksheet.sheet_data[154][2].value.to_s]

# changing cases not to match what is in Spreadsheet(2014)
category = "Unshaded Annual Transmitted Cases 620 and 600"
puts "Populating #{category}"
worksheet.sheet_data[165][2].change_contents(csv_hash['620'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
worksheet.sheet_data[162][2].change_contents(csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
historical_rows << ["#{category} #{worksheet.sheet_data[165][1].value.to_s}",worksheet.sheet_data[165][2].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[162][1].value.to_s}",worksheet.sheet_data[162][2].value.to_s]

# new Unshaded outputs for Std 140 2020
category = "Unshaded Annual Transmitted Cases 660 and 670"
puts "Populating #{category}"
# TODO - add code to populate this
worksheet.sheet_data[163][2].change_contents(csv_hash['660'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
worksheet.sheet_data[164][2].change_contents(csv_hash['670'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
historical_rows << ["#{category} #{worksheet.sheet_data[163][1].value.to_s}",worksheet.sheet_data[163][2].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[164][1].value.to_s}",worksheet.sheet_data[164][2].value.to_s]

category = "Shaded Annual Transmitted Cases 930 and 910"
puts "Populating #{category}"
worksheet.sheet_data[170][2].change_contents(csv_hash['930'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
worksheet.sheet_data[169][2].change_contents(csv_hash['910'][:bestest_building_thermal_envelope_and_fabric_load_reportingzone_total_transmitted_solar_radiation])
historical_rows << ["#{category} #{worksheet.sheet_data[170][1].value.to_s}",worksheet.sheet_data[170][2].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[169][1].value.to_s}",worksheet.sheet_data[169][2].value.to_s]

category = "Sky Temperature Output"
puts "Populating #{category}"
worksheet.sheet_data[177][2].change_contents(csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsky_avg_temp])
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}",worksheet.sheet_data[177][2].value.to_s]

worksheet.sheet_data[177][3].change_contents(csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsky_min_temp])
index_position = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsky_min_index_position]
date_time_array = return_date_time_from_8760_index(index_position)
worksheet.sheet_data[177][4].change_contents(date_time_array[2])
worksheet.sheet_data[177][5].change_contents(date_time_array[3])
worksheet.sheet_data[177][6].change_contents(date_time_array[1])
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}_#{worksheet.sheet_data[176][3].value.to_s}",worksheet.sheet_data[177][3].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}_DATE",date_time_array[0]]
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}_#{worksheet.sheet_data[176][6].value.to_s}",worksheet.sheet_data[177][6].value.to_s]

worksheet.sheet_data[177][7].change_contents(csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsky_max_temp])
index_position = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsky_max_index_position]
date_time_array = return_date_time_from_8760_index(index_position)
worksheet.sheet_data[177][8].change_contents(date_time_array[2])
worksheet.sheet_data[177][9].change_contents(date_time_array[3])
worksheet.sheet_data[177][10].change_contents(date_time_array[1])
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}_#{worksheet.sheet_data[176][7].value.to_s}",worksheet.sheet_data[177][7].value.to_s]
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}_DATE",date_time_array[0]]
historical_rows << ["#{category} #{worksheet.sheet_data[177][1].value.to_s}_#{worksheet.sheet_data[176][10].value.to_s}",worksheet.sheet_data[177][10].value.to_s]

# add Monthly Conditioned Zone Loads (Cases 600 and 900)  
# Reporting measure will need to be updated to get correct data in to workflow_results_csv
category = "Monthly Conditioned Zone Loads Case 600 and 900"
puts "Populating #{category}"

# setup monthly column arrays
array_02 = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_htg_cons].split(",")
array_03 = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_clg_cons].split(",")
array_04 = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_htg_peak_val].split(",")
array_05_06 = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_htg_peak_time].split(",")
array_07 = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_clg_peak_val].split(",")
array_08_09 = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_clg_peak_time].split(",")

array_10 = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_htg_cons].split(",")
array_11 = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_clg_cons].split(",")
array_12 = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_htg_peak_val].split(",")
array_13_14 = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_htg_peak_time].split(",")
array_15 = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_clg_peak_val].split(",")
array_16_17 = csv_hash['900'][:bestest_building_thermal_envelope_and_fabric_load_reportingmonthly_clg_peak_time].split(",")

# populate months with arrays
counter = 0
(189..200).each do |i|

  # total heating 600 aand 900
  worksheet.sheet_data[i][2].change_contents(array_02[counter])
  puts "hello I should be adding a value of #{array_02[counter].to_f}"
  historical_rows << ["#{category} Total Heating 600",worksheet.sheet_data[i][2].value.to_s]
  worksheet.sheet_data[i][10].change_contents(array_10[counter].to_f)
  historical_rows << ["#{category} Total Heating 900",worksheet.sheet_data[i][10].value.to_s]

  # total coooling 600 aand 900
  worksheet.sheet_data[i][3].change_contents(array_03[counter].to_f)
  historical_rows << ["#{category} Total Cooling 600",worksheet.sheet_data[i][3].value.to_s]
  worksheet.sheet_data[i][11].change_contents(array_11[counter].to_f)
  historical_rows << ["#{category} Total Cooling 900",worksheet.sheet_data[i][11].value.to_s]

  # peak heating value 600 aand 900
  worksheet.sheet_data[i][4].change_contents(array_04[counter].to_f)
  historical_rows << ["#{category} Peak Heating Value 600",worksheet.sheet_data[i][4].value.to_s]
  worksheet.sheet_data[i][12].change_contents(array_12[counter].to_f)
  historical_rows << ["#{category} Peak Heating Value 900",worksheet.sheet_data[i][12].value.to_s]

  # TODO - peak heating day time  600 aand 900
  dd_htg_peak_600 = return_date_time_from_string(array_05_06[counter])['dd']
  hh_htg_peak_600 = return_date_time_from_string(array_05_06[counter])['hh']
  dd_htg_peak_900 = return_date_time_from_string(array_13_14[counter])['dd']
  hh_htg_peak_900 = return_date_time_from_string(array_13_14[counter])['hh']
  worksheet.sheet_data[i][5].change_contents(dd_htg_peak_600)
  worksheet.sheet_data[i][6].change_contents(hh_htg_peak_600)
  worksheet.sheet_data[i][13].change_contents(dd_htg_peak_900)
  worksheet.sheet_data[i][14].change_contents(hh_htg_peak_900)
  historical_rows << ["#{category} Peak Heating Day 600",worksheet.sheet_data[i][5].value.to_s]
  historical_rows << ["#{category} Peak Heating Hour 600",worksheet.sheet_data[i][6].value.to_s]
  historical_rows << ["#{category} Peak Heating Day 900",worksheet.sheet_data[i][13].value.to_s]
  historical_rows << ["#{category} Peak Heating Hour 900",worksheet.sheet_data[i][14].value.to_s]

  # peak cooling value 600 aand 900
  worksheet.sheet_data[i][7].change_contents(array_07[counter].to_f)
  historical_rows << ["#{category} Peak Cooling Value 600",worksheet.sheet_data[i][7].value.to_s]
  worksheet.sheet_data[i][15].change_contents(array_15[counter].to_f)
  historical_rows << ["#{category} Peak Cooling Value 900",worksheet.sheet_data[i][15].value.to_s]

  # peak cooling day time  600 aand 900
  dd_clg_peak_600 = return_date_time_from_string(array_08_09[counter])['dd']
  hh_clg_peak_600 = return_date_time_from_string(array_08_09[counter])['hh']
  dd_clg_peak_900 = return_date_time_from_string(array_16_17[counter])['dd']
  hh_clg_peak_900 = return_date_time_from_string(array_16_17[counter])['hh']
  worksheet.sheet_data[i][8].change_contents(dd_clg_peak_600)
  worksheet.sheet_data[i][9].change_contents(hh_clg_peak_600)
  worksheet.sheet_data[i][16].change_contents(dd_clg_peak_900)
  worksheet.sheet_data[i][17].change_contents(hh_clg_peak_900)
  historical_rows << ["#{category} Peak Cooling Day 600",worksheet.sheet_data[i][8].value.to_s]
  historical_rows << ["#{category} Peak Cooling Hour 600",worksheet.sheet_data[i][9].value.to_s]
  historical_rows << ["#{category} Peak Cooling Day 900",worksheet.sheet_data[i][16].value.to_s]
  historical_rows << ["#{category} Peak Cooling Hour 900",worksheet.sheet_data[i][17].value.to_s]

  counter += 1
end

category = "Hourly Incident Solar Radiation Cloudy Day May 4th Case 600 - Horizontal"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0504_zone_surface_roof].split(",")
counter = 0
(229..252).each do |i|
  worksheet.sheet_data[i][2].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Cloudy Day May 4th Case 600 - South"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0504_zone_surface_south].split(",")
counter = 0
(229..252).each do |i|
  worksheet.sheet_data[i][3].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][3].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Cloudy Day May 4th Case 600 - West"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0504_zone_surface_west].split(",")
counter = 0
(229..252).each do |i|
  worksheet.sheet_data[i][4].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][4].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Cloudy Day July 14th Case 600 - Horizontal"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0714_zone_surface_roof].split(",")
counter = 0
(229..252).each do |i|
  worksheet.sheet_data[i][5].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][5].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Clear Day July 14th Case 600 - South"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0714_zone_surface_south].split(",")
counter = 0
(229..252).each do |i|
  worksheet.sheet_data[i][6].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][6].value.to_s]
  counter += 1
end

category = "Hourly Incident Solar Radiation Clear Day July 14th Case 600 - West"
puts "Populating #{category}"
array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_out_inst_slr_rad_0714_zone_surface_west].split(",")
counter = 0
(229..252).each do |i|
  worksheet.sheet_data[i][7].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][7].value.to_s]
  counter += 1
end

# add hourly sky temperatures
category = "Hourly Sky Temperatures for February 1st, May 4th, July 14th Multiple Test Caes 600"
puts "Populating #{category}"
column_target = 8
tot_trans_dates = ['0201','0504','0714']
tot_trans_dates.each do |k|
  if k == '0201'
    array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsite_sky_temp_0201].split(",")
  elsif k == '0504'
    array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsite_sky_temp_0504].split(",")
  elsif k == '0714'
    array = csv_hash['600'][:bestest_building_thermal_envelope_and_fabric_load_reportingsite_sky_temp_0714].split(",")
  else
    array = nil # should not get to this
  end

  counter = 2
  (229..252).each do |i|
    worksheet.sheet_data[i][column_target].change_contents(array[counter].to_f)
    historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][column_target].value.to_s]
    counter += 1
  end
  column_target += 1

end

# add south hourly total transmitted solar for multiple dates and test cases
category = "Hourly Transmitted Solar Radiation per Area February 1st, May 4th, July 14th Multiple Test Cases"
puts "Populating #{category}"
column_target = 11
tot_trans_dates = ['0201','0504','0714']
tot_trans_dates.each do |k|

  tot_trans_test_cases = ['600','660','670']
  tot_trans_test_cases.each do |j|
    
    if k == '0201'
      array = csv_hash[j][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_win_trans_rad_0201_zone_surface_south].split(",")
    elsif k == '0504'
      array = csv_hash[j][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_win_trans_rad_0504_zone_surface_south].split(",")
    elsif k == '0714'
      array = csv_hash[j][:bestest_building_thermal_envelope_and_fabric_load_reportingsurf_win_trans_rad_0714_zone_surface_south].split(",")
    else
      array = nil # should not get to this
    end

    counter = 2
    (229..252).each do |i|
      worksheet.sheet_data[i][column_target].change_contents(array[counter].to_f)
      historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][column_target].value.to_s]
      counter += 1
    end
    column_target += 1

  end  
end

category = "Hourly Heating and Cooling Load February 1st and July 14th Multiple Test Cases"
puts "Populating #{category}"
column_target = 2
(2..25).each do |j| # step through columns in table
  test_case_str = worksheet.sheet_data[259][j].value.to_s
  array_0201 = csv_hash[test_case_str][:bestest_building_thermal_envelope_and_fabric_load_reportingsens_htg_clg_0201].split(",")
  array_0201_c = csv_hash[test_case_str][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0201].split(",")
  array_0714 = csv_hash[test_case_str][:bestest_building_thermal_envelope_and_fabric_load_reportingsens_htg_clg_0714].split(",")
  counter = 2
  (261..284).each do |i|
    if j > 23 # 0201 temp
      worksheet.sheet_data[i][column_target].change_contents(array_0201_c[counter].to_f)
    elsif j > 13 # 0714 kWh
      worksheet.sheet_data[i][column_target].change_contents(array_0714[counter].to_f)
    else # 0201 kWh
      worksheet.sheet_data[i][column_target].change_contents(array_0201[counter].to_f)
    end
    historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][column_target].value.to_s]
    counter += 1
  end
  column_target += 1
end

category = "Hourly Free Float Temperatures February 1st and July 14th Multiple Test Cases"
puts "Populating #{category}"
column_target = 2
(2..7).each do |j| # step through columns in table
  test_case_str = worksheet.sheet_data[291][j].value.to_s
  array_0201 = csv_hash[test_case_str][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0201].split(",")
  array_0714 = csv_hash[test_case_str][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_0714].split(",")
  counter = 2
  (293..316).each do |i|
    if j > 3 && j < 6
      worksheet.sheet_data[i][column_target].change_contents(array_0714[counter].to_f)
    else # 0201 temp
      worksheet.sheet_data[i][column_target].change_contents(array_0201[counter].to_f)
    end
    historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][column_target].value.to_s]
    counter += 1
  end
  column_target += 1
end

category = "Hourly Annual Zone Temperature Bin Data - Case 900FF"
puts "Populating #{category}"
array = csv_hash['900FF'][:bestest_building_thermal_envelope_and_fabric_load_reportingtemp_bins].split(",")
# bin array is just -19 to 70C. The spreadsheet looks for -50 to 98C. May need to extend array or make blanks 0.
# TODO - confirms bins are not offset
counter = 0
(360..449).each do |i|
  worksheet.sheet_data[i][2].change_contents(array[counter+2].to_f)
  historical_rows << ["#{category} #{worksheet.sheet_data[i][1].value.to_s}",worksheet.sheet_data[i][2].value.to_s]
  counter += 1
end

puts "Adding General Information"
# gather general information
common_info = BestestResults.populate_common_info

# starting position
gen_info_row = 21
gen_info_col = 2

# populate general info
worksheet.sheet_data[gen_info_row][gen_info_col].change_contents(common_info[:program_name_and_version])
worksheet.sheet_data[gen_info_row+1][gen_info_col+4].change_contents(common_info[:program_version_release_date])
worksheet.sheet_data[gen_info_row+2][gen_info_col+4].change_contents(common_info[:program_name_short])
worksheet.sheet_data[gen_info_row+3][gen_info_col+4].change_contents(common_info[:results_submission_date])
# row skipped in Excel
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

# load Adding Results worksheet to enter common data
worksheet_ar = workbook['Adding Results']
puts "Loading #{worksheet.sheet_name} Worksheet"

# create OpenStudio copy with updated program info
# Copy Excel File
os_copy_results_5_2a = 'RESULTS5-2A_OS.xlsx'
puts "Making an OpenStudio copy of #{copy_results_5_2a}"
FileUtils.cp(copy_results_5_2a, os_copy_results_5_2a)

puts "Adding General Information"
# gather general information
common_info = BestestResults.populate_common_info("OS")

# starting position
gen_info_row = 21
gen_info_col = 2

# populate general info
worksheet_ar.sheet_data[gen_info_row][gen_info_col].change_contents(common_info[:program_name_and_version])
worksheet_ar.sheet_data[gen_info_row+1][gen_info_col+4].change_contents(common_info[:program_version_release_date])
worksheet_ar.sheet_data[gen_info_row+2][gen_info_col+4].change_contents(common_info[:program_name_short])
worksheet_ar.sheet_data[gen_info_row+3][gen_info_col+4].change_contents(common_info[:results_submission_date])
# row skipped in Excel
worksheet_ar.sheet_data[gen_info_row+5][gen_info_col].change_contents(common_info[:organization])
worksheet_ar.sheet_data[gen_info_row+6][gen_info_col+4].change_contents(common_info[:organization_short])

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