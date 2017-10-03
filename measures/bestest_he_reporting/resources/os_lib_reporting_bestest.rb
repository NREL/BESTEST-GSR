require 'json'

module OsLib_Reporting_Bestest

  # setup - get model, sql, and setup web assets path
  def self.setup(runner)
    results = {}

    # get the last model
    model = runner.lastOpenStudioModel
    if model.empty?
      runner.registerError('Cannot find last model.')
      return false
    end
    model = model.get

    # get the last idf
    workspace = runner.lastEnergyPlusWorkspace
    if workspace.empty?
      runner.registerError('Cannot find last idf file.')
      return false
    end
    workspace = workspace.get

    # get the last sql file
    sqlFile = runner.lastEnergyPlusSqlFile
    if sqlFile.empty?
      runner.registerError('Cannot find last sql file.')
      return false
    end
    sqlFile = sqlFile.get
    model.setSqlFile(sqlFile)

    # populate hash to pass to measure
    results[:model] = model
    # results[:workspace] = workspace
    results[:sqlFile] = sqlFile
    results[:web_asset_path] = OpenStudio.getSharedResourcesPath / OpenStudio::Path.new('web_assets')

    return results
  end

  def self.ann_env_pd(sqlFile)
    # get the weather file run period (as opposed to design day run period)
    ann_env_pd = nil
    sqlFile.availableEnvPeriods.each do |env_pd|
      env_type = sqlFile.environmentType(env_pd)
      if env_type.is_initialized
        if env_type.get == OpenStudio::EnvironmentType.new('WeatherRunPeriod')
          ann_env_pd = env_pd
        end
      end
    end

    return ann_env_pd
  end

  # developer notes
  # - Other thant the 'setup' section above this file should contain methods (def) that create sections and or tables.
  # - Any method that has 'section' in the name will be assumed to define a report section and will automatically be
  # added to the table of contents in the report.
  # - Any section method should have a 'name_only' argument and should stop the method if this is false after the
  # section is defined.
  # - Generally methods that make tables should end with '_table' however this isn't critical. What is important is that
  # it doesn't contain 'section' in the name if it doesn't return a section to the measure.
  # - The data below would typically come from the model or simulation results, but can also come from elsewhere or be
  # defeined in the method as was done with these examples.
  # - You can loop through objects to make a table for each item of that type, such as air loops

  # section for general building information
  def self.general_building_information_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    tables = []

    # gather data for section
    @mat_prop = {}
    @mat_prop[:title] = 'General Building Information'
    @mat_prop[:tables] = tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @mat_prop
    end

    # using helper method that generates table for second example
    tables << OsLib_Reporting_Bestest.general_building_information_table(model, sqlFile, runner)

    return @mat_prop
  end

  # create table with general building information
  # this table shows how to pull information out of the model and the sql file
  def self.general_building_information_table(model, sqlFile, runner)
    # general building information type data output
    general_building_information = {}
    general_building_information[:title] = 'Building Summary' # name will be with section
    general_building_information[:header] = %w(Information Value Units)
    general_building_information[:units] = [] # won't populate for this table since each row has different units.
    general_building_information[:data] = []

    # structure ID / building name
    display = 'Building Name'
    target_units = 'building_name'
    value = model.getBuilding.name.to_s
    general_building_information[:data] << [display, value, target_units]
    runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)

    # net site energy
    display = 'Net Site Energy'
    source_units = 'GJ'
    target_units = 'kBtu'
    value = OpenStudio.convert(sqlFile.netSiteEnergy.get, source_units, target_units).get
    value_neat = OpenStudio.toNeatString(value, 0, true)
    general_building_information[:data] << [display, value_neat, target_units]
    runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)

    # total building area
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='AnnualBuildingUtilityPerformanceSummary' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Building Area' and "
    query << "RowName='Total Building Area' and "
    query << "ColumnName='Area' and "
    query << "Units='m2';"
    query_results = sqlFile.execAndReturnFirstDouble(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for total building area.')
      return false
    else
      display = 'Total Building Area'
      source_units = 'm^2'
      target_units = 'ft^2'
      value = OpenStudio.convert(query_results.get, source_units, target_units).get
      value_neat = OpenStudio.toNeatString(value, 0, true)
      general_building_information[:data] << [display, value_neat, target_units]
      runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)
    end

    # temp code to check OS vs. E+ area
    energy_plus_area = query_results.get
    open_studio_area = model.getBuilding.floorArea
    if not energy_plus_area == open_studio_area
      runner.registerWarning("EnergyPlus reported area is #{query_results.get} (m^2). OpenStudio reported area is #{model.getBuilding.floorArea} (m^2).")
    end

    # EUI
    eui =  sqlFile.netSiteEnergy.get / query_results.get
    display = 'EUI'
    source_units = 'GJ/m^2'
    target_units = 'kBtu/ft^2'
    if query_results.get > 0.0 # don't calculate EUI if building doesn't have any area
      value = OpenStudio.convert(eui, source_units, target_units).get
      value_neat = OpenStudio.toNeatString(value, 4, true)
      runner.registerValue(display.downcase.gsub(" ","_"), value, target_units) # is it ok not to calc EUI if no area in model
    else
      value_neat = "can't calculate EUI."
    end
    general_building_information[:data] << ["#{display} (Based on Net Site Energy and Total Building Area)", value_neat, target_units]

    return general_building_information
  end

  def self.process_output_timeseries (sqlFile, runner, ann_env_pd, time_step, variable_name, key_value)

    output_timeseries = sqlFile.timeSeries(ann_env_pd, time_step, variable_name, key_value)
    if output_timeseries.empty?
      runner.registerWarning("Timeseries not found for #{variable_name}.")
      return false
    else
      runner.registerInfo("Found timeseries for #{variable_name}.")
      output_values = output_timeseries.get.values
      output_times = output_timeseries.get.dateTimes
      array = []
      sum = 0.0
      min = nil
      min_date_time = nil
      max = nil
      max_date_time = nil

      for i in 0..(output_values.size - 1)

        # using this to get average
        array << output_values[i]
        sum += output_values[i]

        # code for min and max
        if min.nil? || output_values[i] < min
          min = output_values[i]
        end
        if max.nil? || output_values[i] > max
          max = output_values[i]
        end

      end
      return {:array => array, :sum => sum, :avg => sum/array.size.to_f, :min => min, :max => max, :min_time => min_date_time, :max_date_time => max_date_time}
    end

  end

  # create output_6_2_1_1_section
  def self.output_6_2_1_1_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    output_6_2_1_1_tables = []

    # gather data for section
    @output_6_2_1_1_section = {}
    @output_6_2_1_1_section[:title] = 'Section 6.2.1.1 All Non-Free-Float Cases'
    @output_6_2_1_1_section[:tables] = output_6_2_1_1_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @output_6_2_1_1_section
    end

    # create table
    table_01 = {}
    table_01[:title] = 'Annual and Peak Heating And Sensible Cooling'
    table_01[:header] = ['Type','Annual Consumption','Peak Value','Peak Time']
    table_01[:units] = ['','MWh','kW']
    table_01[:data] = []

    # annual heating
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='EnergyMeters' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Annual and Peak Values - Other' and "
    query << "RowName='Heating:EnergyTransfer' and "
    query << "ColumnName='Annual Value' and "
    query << "Units='GJ';"
    query_results = sqlFile.execAndReturnFirstDouble(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for heating end use.')
      return false
    else
      display = 'Annual Heating'
      source_units = 'GJ'
      target_units = 'MWh'
      value = OpenStudio.convert(query_results.get, source_units, target_units).get
      annual_htg_value_neat = OpenStudio.toNeatString(value, 4, true)
      runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)
    end

    # annual cooling
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='EnergyMeters' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Annual and Peak Values - Other' and "
    query << "RowName='Cooling:EnergyTransfer' and "
    query << "ColumnName='Annual Value' and "
    query << "Units='GJ';"
    query_results = sqlFile.execAndReturnFirstDouble(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for cooling end use.')
      return false
    else
      display = 'Annual Cooling'
      source_units = 'GJ'
      target_units = 'MWh'
      value = OpenStudio.convert(query_results.get, source_units, target_units).get
      annual_clg_value_neat = OpenStudio.toNeatString(value, 4, true)
      runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)
    end

    # peak heating
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='EnergyMeters' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Annual and Peak Values - Other' and "
    query << "RowName='Heating:EnergyTransfer' and "
    query << "ColumnName='Maximum Value' and "
    query << "Units='W';"
    query_results = sqlFile.execAndReturnFirstDouble(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for heating peak value.')
      return false
    else
      display = 'Peak Heating Value'
      source_units = 'W'
      target_units = 'kW'
      #value = OpenStudio.convert(query_results.get, source_units, target_units).get
      #peak_htg_value_neat = OpenStudio.toNeatString(value, 4, true)
      #runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)

      # change peak heating to look at max hourly consumption.
      hourly_heating_kwh = OsLib_Reporting_Bestest.hourly_heating_peak(model, sqlFile, runner)
      runner.registerValue(display.downcase.gsub(" ","_"), hourly_heating_kwh.max, target_units)

    end

    # peak heating time
    # todo - update peak cooling time to use hourly consumption instead of tabular results
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='EnergyMeters' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Annual and Peak Values - Other' and "
    query << "RowName='Heating:EnergyTransfer' and "
    query << "ColumnName='Timestamp of Maximum {TIMESTAMP}' and "
    query << "Units='';"
    query_results = sqlFile.execAndReturnFirstString(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for heating peak timestep.')
      return false
    else
      display = 'Peak Heating Time'
      peak_htg_time = query_results.get
      runner.registerValue(display.downcase.gsub(" ","_"), peak_htg_time)
    end


    # peak cooling
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='EnergyMeters' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Annual and Peak Values - Other' and "
    query << "RowName='Cooling:EnergyTransfer' and "
    query << "ColumnName='Maximum Value' and "
    query << "Units='W';"
    query_results = sqlFile.execAndReturnFirstDouble(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for cooling peak value.')
      return false
    else
      display = 'Peak Cooling Value'
      source_units = 'W'
      target_units = 'kW'
      #value = OpenStudio.convert(query_results.get, source_units, target_units).get
      #peak_clg_value_neat = OpenStudio.toNeatString(value, 4, true)
      #runner.registerValue(display.downcase.gsub(" ","_"), value, target_units)

      # change peak heating to look at max hourly consumption
      hourly_cooling_kwh = OsLib_Reporting_Bestest.hourly_cooling_peak(model, sqlFile, runner)
      runner.registerValue(display.downcase.gsub(" ","_"), hourly_cooling_kwh.max, target_units)

    end

    # peak cooling time
    # todo - update peak cooling time to use hourly consumption instead of tabular results
    query = 'SELECT Value FROM tabulardatawithstrings WHERE '
    query << "ReportName='EnergyMeters' and "
    query << "ReportForString='Entire Facility' and "
    query << "TableName='Annual and Peak Values - Other' and "
    query << "RowName='Cooling:EnergyTransfer' and "
    query << "ColumnName='Timestamp of Maximum {TIMESTAMP}' and "
    query << "Units='';"
    query_results = sqlFile.execAndReturnFirstString(query)
    if query_results.empty?
      runner.registerWarning('Did not find value for heating peak timestep.')
      return false
    else
      display = 'Peak Cooling Time'
      peak_clg_time = query_results.get
      runner.registerValue(display.downcase.gsub(" ","_"), peak_clg_time)
    end

    # add rows to table
    table_01[:data] << ['Heating', annual_htg_value_neat,peak_htg_value_neat,peak_htg_time]
    table_01[:data] << ['Cooling', annual_clg_value_neat,peak_clg_value_neat,peak_clg_time]

    # add table to array of tables
    output_6_2_1_1_tables << table_01

    return @output_6_2_1_1_section
  end

  # create table_6_1_section
  def self.table_6_1_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    table_6_1_tables = []

    # gather data for section
    @table_6_1_section = {}
    @table_6_1_section[:title] = 'Table 6-1 Daily Hourly Output Requirements'
    @table_6_1_section[:tables] = table_6_1_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @table_6_1_section
    end

    # create table
    table_01 = {}
    table_01[:title] = 'Hourly Incident Unshaded Solar Radiation (W/m^2)'
    table_01[:header] = ['Date','Orientation',1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]
    table_01[:units] = [] # list units in title vs. in each column
    table_01[:data] = []

    # get time series data for main zone
    ann_env_pd = OsLib_Reporting_Bestest.ann_env_pd(sqlFile)
    if ann_env_pd
      # get keys
      keys = sqlFile.availableKeyValues(ann_env_pd, 'Hourly', 'Surface Outside Face Incident Solar Radiation Rate per Area')

      if keys.include? 'ZONE ONE'
        key = 'ZONE ONE'
      elsif keys.include? 'SUN ZONE'
        key = 'SUN ZONE'
      end

      # todo - should it be Wh/m^2
      source_units = 'W/m^2'
      target_units = 'W/m^2'

      # loop through surfaces
      model.getSurfaces.each do |surface|
        next if OpenStudio::convert(surface.tilt,"rad","deg").get.round == 0
        next if OpenStudio::convert(surface.azimuth,"rad","deg").get.round == 0
        next if OpenStudio::convert(surface.azimuth,"rad","deg").get.round == 90
        key = surface.name.to_s.upcase

        # get values
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Surface Outside Face Incident Solar Radiation Rate per Area', key)
        if output_timeseries.is_initialized # checks to see if time_series exists

          # get March 5th values
          row_data = ['March 5',surface.name.to_s.upcase]
          table_01[:header].each do |hour|
            next if hour == "Date"
            next if hour == "Orientation"
            date_string = "2009-03-05 #{hour}:00:00.000"
            date_time = OpenStudio::DateTime.new(date_string)
            val_at_date_time = output_timeseries.get.value(date_time)
            value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
            row_data << value.round(2)
          end
          runner.registerValue("surf_out_inst_slr_rad_0305_#{surface.name.get.downcase.gsub(" ","_")}",row_data.to_s)
          table_01[:data] << row_data

          # get July 27th values
          row_data = ['July 22',surface.name.to_s.upcase]
          table_01[:header].each do |hour|
            next if hour == "Date"
            next if hour == "Orientation"
            date_string = "2009-07-27 #{hour}:00:00.000"
            date_time = OpenStudio::DateTime.new(date_string)
            val_at_date_time = output_timeseries.get.value(date_time)
            value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
            row_data << value.round(2)
          end
          runner.registerValue("surf_out_inst_slr_rad_0727_#{surface.name.get.downcase.gsub(" ","_")}",row_data.to_s)
          table_01[:data] << row_data

        else
          runner.registerWarning("Didn't find data for Outside Face Incident Solar Radiation Rate per Area")
        end # end of if output_timeseries.is_initialized

      end
    end

    # add table to array of tables
    if table_01[:data].size > 0
      table_6_1_tables << table_01
    end

    # use helper method that generates additional table for section
    table_6_1_tables << OsLib_Reporting_Bestest.hourly_heating_cooling_table(model, sqlFile, runner)
    table_6_1_tables << OsLib_Reporting_Bestest.free_floating_temp(model, sqlFile, runner)

    return @table_6_1_section
  end

  # create hourly_heating_cooling_table
  def self.hourly_heating_cooling_table(model, sqlFile, runner)

    # FF case gives ruby error on server for this but not local. This should skip it to avoid server
    if model.getBuilding.name.get.include?("FF")
      return nil
    end

    table = {}
    table[:title] = 'Hourly Loads (kWh)'
    table[:header] = ['Date','Type',1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]
    table[:units] = [] # list units in title vs. in each column
    table[:data] = []

    # get time series data for main zone
    ann_env_pd = OsLib_Reporting_Bestest.ann_env_pd(sqlFile)
    if ann_env_pd
      # get keys
      keys = sqlFile.availableKeyValues(ann_env_pd, 'Hourly', 'Zone Air System Sensible Heating Energy')

      if keys.include? 'ZONE ONE'
        key = 'ZONE ONE'
      elsif keys.include? 'BACK ZONE'
        key = 'BACK ZONE'
      end

      source_units = 'J'
      target_units = 'kWh'
      hourly_htg = nil
      hourly_clg = nil

      # get heating values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Air System Sensible Heating Energy', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        # get January 4th values
        row_data = ['January 4','Heating']
        table[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Type"
          date_string = "2009-01-04 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
          row_data << value.round(2)
        end
        hourly_htg = row_data
        table[:data] << row_data

      else
        runner.registerWarning("Didn't find data for Zone Air System Sensible Heating Energy")
      end # end of if output_timeseries.is_initialized

      # get heating values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Air System Sensible Cooling Energy', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        # get January 4th values
        row_data = ['January 4','Cooling']
        table[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Type"
          date_string = "2009-01-04 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
          row_data << value.round(2)
        end
        hourly_clg = row_data
        table[:data] << row_data

      else
        runner.registerWarning("Didn't find data for Zone Air System Sensible Cooling Energy")
      end # end of if output_timeseries.is_initialized

      # combine headting and cooling into one array
      combined_hourly = []
      26.times do |i|
        if i < 2
          combined_hourly << hourly_htg[i] + hourly_clg[i]
        else
          combined_hourly << hourly_htg[i] - hourly_clg[i]
        end
      end
      runner.registerValue("sens_htg_clg_0104",combined_hourly.to_s)


    end

    return table
  end

  # get peak hourly heating value (maximum hourly consumption)
  def self.hourly_heating_peak(model, sqlFile, runner)

    # FF case gives ruby error on server for this but not local. This should skip it to avoid server
    if model.getBuilding.name.get.include?("FF")
      return nil
    end

    array_8760 = [] # values

    ann_env_pd = OsLib_Reporting_Bestest.ann_env_pd(sqlFile)
    if ann_env_pd
      # get keys
      keys = sqlFile.availableKeyValues(ann_env_pd, 'Hourly', 'Zone Air System Sensible Heating Energy')

      if keys.include? 'ZONE ONE'
        key = 'ZONE ONE'
      elsif keys.include? 'BACK ZONE'
        key = 'BACK ZONE'
      end

      source_units = 'J'
      target_units = 'kWh'

      # get heating values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Air System Sensible Heating Energy', key.to_s)
      if output_timeseries.is_initialized # checks to see if time_series exists

        output_timeseries = output_timeseries.get.values
        for i in 0..(output_timeseries.size - 1)
          value = OpenStudio.convert(output_timeseries[i], source_units, target_units).get
          array_8760 << value
        end

      end
    end

    return array_8760

  end

  # get peak hourly cooling value (maximum hourly consumption)
  def self.hourly_cooling_peak(model, sqlFile, runner)

    # FF case gives ruby error on server for this but not local. This should skip it to avoid server
    if model.getBuilding.name.get.include?("FF")
      return nil
    end

    array_8760 = []

    ann_env_pd = OsLib_Reporting_Bestest.ann_env_pd(sqlFile)
    if ann_env_pd
      # get keys
      keys = sqlFile.availableKeyValues(ann_env_pd, 'Hourly', 'Zone Air System Sensible Cooling Energy')

      if keys.include? 'ZONE ONE'
        key = 'ZONE ONE'
      elsif keys.include? 'BACK ZONE'
        key = 'BACK ZONE'
      end

      source_units = 'J'
      target_units = 'kWh'

      # get heating values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Air System Sensible Cooling Energy', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        output_timeseries = output_timeseries.get.values
        for i in 0..(output_timeseries.size - 1)
          value = OpenStudio.convert(output_timeseries[i], source_units, target_units).get
          array_8760 << value
        end

      end
    end

    return array_8760

  end

  # create free_floating_temp
  def self.free_floating_temp(model, sqlFile, runner)
    table = {}
    table[:title] = 'Hourly Zone Mean Air Temperature (C)' # only show for Free-Floating cases
    table[:header] = ['Date',1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]
    table[:units] = [] # list units in title vs. in each column
    table[:data] = []


    # get time series data for main zone
    ann_env_pd = OsLib_Reporting_Bestest.ann_env_pd(sqlFile)
    if ann_env_pd
      # get keys
      keys = sqlFile.availableKeyValues(ann_env_pd, 'Hourly', 'Zone Mean Air Temperature')

      if keys.include? 'ZONE ONE'
        key = 'ZONE ONE'
      elsif keys.include? 'SUN ZONE'
        key = 'SUN ZONE'
      end

      # create array from values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Mean Air Temperature', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        # get January 4th values
        row_data = ['January 4']
        table[:header].each do |hour|
          next if hour == "Date"
          date_string = "2009-01-04 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(1)
        end
        runner.registerValue("temp_0104",row_data.to_s)
        table[:data] << row_data

        # get July 27th values
        row_data = ['July 27']
        table[:header].each do |hour|
          next if hour == "Date"
          date_string = "2009-07-27 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(1)
        end
        runner.registerValue("temp_0727",row_data.to_s)
        table[:data] << row_data

      else
        runner.registerWarning("Didn't find data for Zone Mean Air Temperature")
      end # end of if output_timeseries.is_initialized

    end

    return table
  end

  # create case_600_only_section
  def self.case_600_only_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    case_600_only_tables = []

    # gather data for section
    @case_600_only_section = {}
    @case_600_only_section[:title] = 'Section 6.2.1.2 Case 600 Only'
    @case_600_only_section[:tables] = case_600_only_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @case_600_only_section
    end

    # create table
    table_01 = {}
    table_01[:title] = "Annual Incident Unshaded Total Solar Radiation (diffuse and direct)"
    table_01[:header] = ['Direction','Radiation']
    table_01[:units] = ['','kWh/m^2']
    table_01[:data] = []

    # TODO -  should update to fail gracefully if this isn't valid
    # get the weather file run period (as opposed to design day run period)
    ann_env_pd = nil
    sqlFile.availableEnvPeriods.each do |env_pd|
      env_type = sqlFile.environmentType(env_pd)
      if env_type.is_initialized
        if env_type.get == OpenStudio::EnvironmentType.new("WeatherRunPeriod")
          ann_env_pd = env_pd
          break
        end
      end
    end

    # north_incident_solar_radiation
    key_value =  "ZONE SURFACE NORTH"
    variable_name = "Surface Outside Face Incident Solar Radiation Rate per Area"
    timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
    value_kwh_m2 = OpenStudio.convert(timeseries_hash[:sum],'Wh/m^2','kWh/m^2').get # using Wh since timestep is hourly
    runner.registerValue('north_incident_solar_radiation',value_kwh_m2,'kWh/m^2')
    table_01[:data] << ['North',value_kwh_m2.round(2)]
    # east_incident_solar_radiation
    key_value =  "ZONE SURFACE EAST"
    variable_name = "Surface Outside Face Incident Solar Radiation Rate per Area"
    timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
    value_kwh_m2 = OpenStudio.convert(timeseries_hash[:sum],'Wh/m^2','kWh/m^2').get # using Wh since timestep is hourly
    runner.registerValue('east_incident_solar_radiation',value_kwh_m2,'kWh/m^2')
    table_01[:data] << ['East',value_kwh_m2.round(2)]
    # west_incident_solar_radiation
    key_value =  "ZONE SURFACE WEST"
    variable_name = "Surface Outside Face Incident Solar Radiation Rate per Area"
    timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
    value_kwh_m2 = OpenStudio.convert(timeseries_hash[:sum],'Wh/m^2','kWh/m^2').get # using Wh since timestep is hourly
    runner.registerValue('west_incident_solar_radiation',value_kwh_m2,'kWh/m^2')
    table_01[:data] << ['West',value_kwh_m2.round(2)]
    # south_incident_solar_radiation
    key_value =  "ZONE SURFACE SOUTH"
    variable_name = "Surface Outside Face Incident Solar Radiation Rate per Area"
    timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
    value_kwh_m2 = OpenStudio.convert(timeseries_hash[:sum],'Wh/m^2','kWh/m^2').get # using Wh since timestep is hourly
    runner.registerValue('south_incident_solar_radiation',value_kwh_m2,'kWh/m^2')
    table_01[:data] << ['South',value_kwh_m2.round(2)]
    # horizontal_incident_solar_radiation
    key_value =  "ZONE SURFACE ROOF"
    variable_name = "Surface Outside Face Incident Solar Radiation Rate per Area"
    timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
    value_kwh_m2 = OpenStudio.convert(timeseries_hash[:sum],'Wh/m^2','kWh/m^2').get # using Wh since timestep is hourly
    runner.registerValue('horizontal_incident_solar_radiation',value_kwh_m2,'kWh/m^2')
    table_01[:data] << ['Horizontal',value_kwh_m2.round(2)]

    # add table to array of tables
    case_600_only_tables << table_01

    # create table
    table_02 = {}
    table_02[:title] = "Unshaded Annual Transmitted Solar Radiation (diffuse and direct) Through South Windows"
    table_02[:header] = ['Direction','Radiation']
    table_02[:units] = ['','kWh/m^2']
    table_02[:data] = []

    # add rows to table
    table_02[:data] << ['South',]

    # add table to array of tables
    # case_600_only_tables << table_02

    return @case_600_only_section
  end

  # create case_9xx_only_section
  def self.case_9xx_only_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    case_9xx_only_tables = []

    # gather data for section
    @case_9xx_only_section = {}
    @case_9xx_only_section[:title] = 'runner.registerValues for 9xx cases'
    @case_9xx_only_section[:tables] = case_9xx_only_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @case_9xx_only_section
    end

    # TODO -  should update to fail gracefully if this isn't valid
    # get the weather file run period (as opposed to design day run period)
    ann_env_pd = nil
    sqlFile.availableEnvPeriods.each do |env_pd|
      env_type = sqlFile.environmentType(env_pd)
      if env_type.is_initialized
        if env_type.get == OpenStudio::EnvironmentType.new("WeatherRunPeriod")
          ann_env_pd = env_pd
          break
        end
      end
    end

    # zone_total_transmitted_beam_solar_radiation
    key_value =  "ZONE ONE"
    variable_name = "Zone Windows Total Transmitted Solar Radiation Rate"
    timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
    value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get # using Wh since timestep is hourly
    value_kwh_m2 = value_kwh / 12.0 # all zones with windows have 12m^2
    runner.registerValue('zone_total_transmitted_solar_radiation',value_kwh_m2,'kWh/m^2')

    return @case_9xx_only_section
  end

  # create case_610_only_section
  def self.case_610_only_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    case_610_only_tables = []

    # gather data for section
    @case_610_only_section = {}
    @case_610_only_section[:title] = 'Section 6.2.1.3 Case 610 Only'
    @case_610_only_section[:tables] = case_610_only_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @case_610_only_section
    end

    # create table
    table_01 = {}
    table_01[:title] = "Annual TransmittedSolar Radiation Through the Shaded South Window with Horizontal Overhang"
    table_01[:header] = ['Direction','Radiation']
    table_01[:units] = ['','kWh/m^2']
    table_01[:data] = []

    # add rows to table
    table_01[:data] << ['South',]

    # add table to array of tables
    case_610_only_tables << table_01

    return @case_610_only_section
  end

  # create case_620_only_section
  def self.case_620_only_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    case_620_only_tables = []

    # gather data for section
    @case_620_only_section = {}
    @case_620_only_section[:title] = 'Section 6.2.1.4 Case 620 Only'
    @case_620_only_section[:tables] = case_620_only_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @case_620_only_section
    end

    # create table
    table_01 = {}
    table_01[:title] = "Unshaded Annual Transmitted Solar Radiation (diffuse and direct) Through West Windows"
    table_01[:header] = ['Direction','Radiation']
    table_01[:units] = ['','kWh/m^2']
    table_01[:data] = []

    # add rows to table
    table_01[:data] << ['West',]

    # add table to array of tables
    case_620_only_tables << table_01

    return @case_620_only_section
  end

  # create case_630_only_section
  def self.case_630_only_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    case_630_only_tables = []

    # gather data for section
    @case_630_only_section = {}
    @case_630_only_section[:title] = 'Section 6.2.1.5 Case 630 Only'
    @case_630_only_section[:tables] = case_630_only_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @case_630_only_section
    end

    # create table
    table_01 = {}
    table_01[:title] = "Annual Transmitted Solar Radiation Through the Shaded West Window with Horizontal Overhang and Vertical Fins"
    table_01[:header] = ['Direction','Radiation']
    table_01[:units] = ['','kWh/m^2']
    table_01[:data] = []

    # add rows to table
    table_01[:data] << ['West',]

    # add table to array of tables
    case_630_only_tables << table_01

    return @case_630_only_section
  end

  # create ff_temp_bins_section
  def self.ff_temp_bins_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    ff_temp_bins_tables = []

    # gather data for section
    @ff_temp_bins_section = {}
    @ff_temp_bins_section[:title] = 'Section 6.2.1.7 Case 900FF Only'
    @ff_temp_bins_section[:tables] = ff_temp_bins_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @ff_temp_bins_section
    end

    # create table
    table_01 = {}
    table_01[:title] = "Hourly Zone Mean Temperature Bins (1C bin size)"
    table_01[:header] = ['Temperature','Bin Size'] # for now do row for each bind
    table_01[:units] = ['C','Hours']
    table_01[:data] = []

    # gather data (we can pre-poplulate 0 value from -20C to 70C if needed)
    hourly_values_rnd = {}
    array_8760 = []

    # get time series data for main zone
    ann_env_pd = OsLib_Reporting_Bestest.ann_env_pd(sqlFile)
    if ann_env_pd
      # get keys
      keys = sqlFile.availableKeyValues(ann_env_pd, 'Hourly', 'Zone Mean Air Temperature')

      if keys.include? 'ZONE ONE'
        key = 'ZONE ONE'
      elsif keys.include? 'SUN ZONE'
        key = 'SUN ZONE'
      end

      # create array from values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Mean Air Temperature', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        output_timeseries = output_timeseries.get.values
        for i in 0..(output_timeseries.size - 1)

          # using this to get average
          array_8760 << output_timeseries[i]

          if output_timeseries[i].truncate != output_timeseries[i] and output_timeseries[i] < 0
            # without this negeative numbers seem to truncate towards zero vs. colder temp
            value_truncate = output_timeseries[i].truncate - 1
          else
            value_truncate = output_timeseries[i].truncate
          end
          if hourly_values_rnd.has_key?(value_truncate)
            hourly_values_rnd[value_truncate] += 1
          else
            hourly_values_rnd[value_truncate] = 1
          end
        end
      else
        runner.registerWarning("Didn't find data for Zone Mean Air Temperature")
      end # end of if output_timeseries.is_initialized

    end

    # add rows to table
    hourly_values_rnd.sort_by { |k,v| k}.each do |k,v|
      table_01[:data] << [k,v]
    end

    # create array from -20C through 70C for register value
    full_temp_bin = []
    (-20..70).each do |i|
      if hourly_values_rnd[i]
        full_temp_bin << hourly_values_rnd[i]
      else
        full_temp_bin << 0
      end
    end
    runner.registerValue("temp_bins",full_temp_bin.to_s)

    # store min and max and avg temps as register value, along with index position
    # will convert index to date/time downstream
    runner.registerValue('min_temp',array_8760.min,'C')
    runner.registerValue('min_index_position',array_8760.each_with_index.min[1])
    runner.registerValue('max_temp',array_8760.max,'C')
    runner.registerValue('max_index_position',array_8760.each_with_index.max[1])
    runner.registerValue('avg_temp',array_8760.reduce(:+) / array_8760.size.to_f,'C')

    # add table to array of tables
    ff_temp_bins_tables << table_01

    return @ff_temp_bins_section
  end

end
