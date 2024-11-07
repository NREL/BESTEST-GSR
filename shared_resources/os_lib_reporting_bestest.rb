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
    bldg_name = workspace.getObjectsByType("Building".to_IddObjectType).first.getString(0).get

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

      # hourly sky temps
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Site Sky Temperature', 'Environment')
      if output_timeseries.is_initialized # checks to see if time_series exists (only case 600)

        # get Febuary 1st values
        row_data = ['February 1','Site Variable']
        table_01[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Orientation"
          date_string = "2009-02-01 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(2)
        end
        runner.registerValue("site_sky_temp_0201",row_data.join(","))

        # get Febuary 1st values
        row_data = ['May 4','Site Variable']
        table_01[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Orientation"
          date_string = "2009-05-04 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(2)
        end
        runner.registerValue("site_sky_temp_0504",row_data.join(","))

        # get Febuary 1st values
        row_data = ['July 14','Site Variable']
        table_01[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Orientation"
          date_string = "2009-7-14 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(2)
        end
        runner.registerValue("site_sky_temp_0714",row_data.join(","))

        # TODO - get extreme and avg annual sky conditions for Sky Temperature Output Table
        sky_t_array_8760 = []
        for i in 0..(output_timeseries.get.values.size - 1)
          # using this to get average
          sky_t_array_8760 << output_timeseries.get.values[i]
        end

        # store min and max and avg temps as register value, along with index position
        # will convert index to date/time downstream
        runner.registerValue('sky_min_temp',sky_t_array_8760.min,'C')
        runner.registerValue('sky_min_index_position',sky_t_array_8760.each_with_index.min[1])
        runner.registerValue('sky_max_temp',sky_t_array_8760.max,'C')
        runner.registerValue('sky_max_index_position',sky_t_array_8760.each_with_index.max[1])
        runner.registerValue('sky_avg_temp',sky_t_array_8760.reduce(:+) / sky_t_array_8760.size.to_f,'C')

      end

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

      # create temp model from workspace and check orientation
      workspace = runner.lastEnergyPlusWorkspace.get
      rt = OpenStudio::EnergyPlus::ReverseTranslator.new
      model2 = rt.translateWorkspace(workspace)

      # loop through surfaces
      model2.getSurfaces.each do |surface|
        next if OpenStudio::convert(surface.azimuth,"rad","deg").get.round == 0 && OpenStudio::convert(surface.tilt,"rad","deg").get.round != 0
        next if OpenStudio::convert(surface.azimuth,"rad","deg").get.round == 90 && OpenStudio::convert(surface.tilt,"rad","deg").get.round != 0
        key = surface.name.to_s.upcase

        # get values
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Surface Outside Face Incident Solar Radiation Rate per Area', key)
        if output_timeseries.is_initialized # checks to see if time_series exists

          # get Febuary 1st values
          row_data = ['February 1',surface.name.to_s.upcase]
          table_01[:header].each do |hour|
            next if hour == "Date"
            next if hour == "Orientation"
            date_string = "2009-02-01 #{hour}:00:00.000"
            date_time = OpenStudio::DateTime.new(date_string)
            val_at_date_time = output_timeseries.get.value(date_time)
            value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
            row_data << value.round(2)
          end
          runner.registerValue("surf_out_inst_slr_rad_0201_#{surface.name.get.downcase.gsub(" ","_")}",row_data.to_s)
          table_01[:data] << row_data

          # get May 4th values
          row_data = ['May 4',surface.name.to_s.upcase]
          table_01[:header].each do |hour|
            next if hour == "Date"
            next if hour == "Orientation"
            date_string = "2009-05-04 #{hour}:00:00.000"
            date_time = OpenStudio::DateTime.new(date_string)
            val_at_date_time = output_timeseries.get.value(date_time)
            value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
            row_data << value.round(2)
          end
          runner.registerValue("surf_out_inst_slr_rad_0504_#{surface.name.get.downcase.gsub(" ","_")}",row_data.to_s)
          table_01[:data] << row_data

          # get July 14th values
          row_data = ['July 14',surface.name.to_s.upcase]
          table_01[:header].each do |hour|
            next if hour == "Date"
            next if hour == "Orientation"
            date_string = "2009-07-14 #{hour}:00:00.000"
            date_time = OpenStudio::DateTime.new(date_string)
            val_at_date_time = output_timeseries.get.value(date_time)
            value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
            row_data << value.round(2)
          end
          runner.registerValue("surf_out_inst_slr_rad_0714_#{surface.name.get.downcase.gsub(" ","_")}",row_data.to_s)
          table_01[:data] << row_data
        else
          runner.registerWarning("Didn't find data for Outside Face Incident Solar Radiation Rate per Area")
        end # end of if output_timeseries.is_initialized
      end
    end

    # loop through surfaces to find south facing windows
    found_south = false
    model2.getSubSurfaces.sort.each do |sub_surface|
      
      # TODO - there are two sub-surafces in same base surface, ideally can sum area and radiation of both vs. picking first
      # but with no overhang they should be the same.
      next if found_south
      next if OpenStudio::convert(sub_surface.azimuth,"rad","deg").get.round != 180
      next if OpenStudio::convert(sub_surface.tilt,"rad","deg").get.round != 90
      key = sub_surface.name.to_s.upcase
      sub_surface_area = sub_surface.grossArea

      # get parent surface to use in registerValue
      parent_surf_name = sub_surface.surface.get.name.get.downcase.gsub(" ","_")

      # get value
      # Transmitted Solar Radiation (W) isn't avaiable as variable per area. Divde by area to get Wh/m^2
      # just adding registerValues, not extending tables. Tables only go to HTML which isn't maintained and may be removed at some point.
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Surface Window Transmitted Solar Radiation Rate', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        # get Febuary 1st values
        row_data = ['February 1',sub_surface.name.to_s.upcase]
        table_01[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Orientation"
          date_string = "2009-02-01 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          val_at_date_time = val_at_date_time/sub_surface_area # get as W/m^2
          row_data << val_at_date_time.round(2)
        end
        runner.registerValue("surf_win_trans_rad_0201_#{parent_surf_name}",row_data.to_s)

        # get May 4th values
        row_data = ['May 4',sub_surface.name.to_s.upcase]
        table_01[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Orientation"
          date_string = "2009-05-04 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          val_at_date_time = val_at_date_time/sub_surface_area # get as W/m^2
          row_data << val_at_date_time.round(2)
        end
        runner.registerValue("surf_win_trans_rad_0504_#{parent_surf_name}",row_data.to_s)

        # get July 14th values
        row_data = ['July 14',sub_surface.name.to_s.upcase]
        table_01[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Orientation"
          date_string = "2009-07-14 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          val_at_date_time = val_at_date_time/sub_surface_area # get as W/m^2          
          row_data << val_at_date_time.round(2)
        end
        runner.registerValue("surf_win_trans_rad_0714_#{parent_surf_name}",row_data.to_s)

        # for now stopping after first south window instead of averaging multiple windows
        found_south = true
      else
        runner.registerWarning("Didn't find data for Surface Window Transmitted Solar Radiation Rate")
      end # end of if output_timeseries.is_initialized
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
    if runner.lastEnergyPlusWorkspace.get.getObjectsByType("Building".to_IddObjectType).first.getString(0).get.include?("FF")
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
      hourly_htg_0201 = nil
      hourly_clg_0201 = nil
      hourly_htg_0714 = nil
      hourly_clg_0714 = nil

      # get heating values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Air System Sensible Heating Energy', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        # get January 4th values
        row_data = ['Febuary 1','Heating']
        table[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Type"
          date_string = "2009-02-01 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
          row_data << value.round(2)
        end
        hourly_htg_0201 = row_data
        table[:data] << row_data

        # get July 14th values
        row_data = ['July 14','Heating']
        table[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Type"
          date_string = "2009-07-14 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
          row_data << value.round(2)
        end
        hourly_htg_0714 = row_data
        table[:data] << row_data

      else
        runner.registerWarning("Didn't find data for Zone Air System Sensible Heating Energy")
      end # end of if output_timeseries.is_initialized

      # get heating values
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', 'Zone Air System Sensible Cooling Energy', key)
      if output_timeseries.is_initialized # checks to see if time_series exists

        # get January 4th values
        row_data = ['Febuary 1','Cooling']
        table[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Type"
          date_string = "2009-02-01 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
          row_data << value.round(2)
        end
        hourly_clg_0201 = row_data
        table[:data] << row_data

        # get January 4th values
        row_data = ['July 14','Cooling']
        table[:header].each do |hour|
          next if hour == "Date"
          next if hour == "Type"
          date_string = "2009-07-014 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_units, target_units).get
          row_data << value.round(2)
        end
        hourly_clg_0714 = row_data
        table[:data] << row_data  

      else
        runner.registerWarning("Didn't find data for Zone Air System Sensible Cooling Energy")
      end # end of if output_timeseries.is_initialized

      # combine headting and cooling into one array
      combined_hourly_0201 = []
      combined_hourly_0714 = []
      26.times do |i|
        if i < 2 # header strings
          combined_hourly_0201 << hourly_htg_0201[i] + hourly_clg_0201[i]
          combined_hourly_0714 << hourly_htg_0714[i] + hourly_clg_0714[i]
        else
          combined_hourly_0201 << hourly_htg_0201[i] - hourly_clg_0201[i]
          combined_hourly_0714 << hourly_htg_0714[i] - hourly_clg_0714[i]
        end
      end

      runner.registerInfo("End of method, Adding values for heating and cooling on February 1st and July 14th.")
      runner.registerValue("sens_htg_clg_0201",combined_hourly_0201.to_s)
      runner.registerValue("sens_htg_clg_0714",combined_hourly_0714.to_s)

    end

    return table
  end

  # get peak hourly heating value (maximum hourly consumption)
  def self.hourly_heating_peak(model, sqlFile, runner)

    # FF case gives ruby error on server for this but not local. This should skip it to avoid server
    if runner.lastEnergyPlusWorkspace.get.getObjectsByType("Building".to_IddObjectType).first.getString(0).get.include?("FF")
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
    if runner.lastEnergyPlusWorkspace.get.getObjectsByType("Building".to_IddObjectType).first.getString(0).get.include?("FF")
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

        # get Febuary 1st values
        row_data = ['Febuary 1']
        table[:header].each do |hour|
          next if hour == "Date"
          date_string = "2009-02-01 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(1)
        end
        runner.registerValue("temp_0201",row_data.to_s)
        table[:data] << row_data

        # get Julyn14th values
        row_data = ['July 14']
        table[:header].each do |hour|
          next if hour == "Date"
          date_string = "2009-07-14 #{hour}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          row_data << val_at_date_time.round(1)
        end
        runner.registerValue("temp_0714",row_data.to_s)
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
    key_value =  "ZONE ONE SPACE"
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

  # create case_610_only_section
  def self.monthly_htg_clg_table_section(model, sqlFile, runner, name_only = false)
    # array to hold tables
    monthly_htg_clg_tables = []

    # gather data for section
    @monthly_htg_clg_section = {}
    @monthly_htg_clg_section[:title] = 'Monthly Heating and Cooling for Casees 600 and 900'
    @monthly_htg_clg_section[:tables] = monthly_htg_clg_tables

    # stop here if only name is requested this is used to populate display name for arguments
    if name_only == true
      return @monthly_htg_clg_section
    end

    # array's for runnerRegisterValues
    monthly_htg_cons = []
    monthly_clg_cons = []
    monthly_htg_peak_val = []
    monthly_htg_peak_time = []
    monthly_clg_peak_val = []
    monthly_clg_peak_time = []

    # only case 600 or 900
    name_test = model.getBuilding.name.get.to_s.gsub('BESTEST Case ','')[0..2].to_s # change logic if FF case added that has more characters
    if !['600','900'].include?(name_test)
      return @monthly_htg_clg_section
    end

    runner.registerInfo("Starting to get monthly heating and cooling data.")
    months = ['January','February','March','April','May','June','July','August','September','October','November','December']
    months.each do |month|

      # keep long month for sql query use short for reg_value so matches Excel labe
      mon = month[0..2].downcase

      # gather monthly heating from BUILDING ENERGY PERFORMANCE - DISTRICT HEATING Custom Monthly Report
      query = 'SELECT Value FROM tabulardatawithstrings WHERE '
      query << "ReportName='BUILDING ENERGY PERFORMANCE - DISTRICT HEATING WATER' and "
      query << "ReportForString='Meter' and "
      query << "TableName='Custom Monthly Report' and "
      query << "RowName='#{month}' and "
      query << "ColumnName='HEATING:DISTRICTHEATINGWATER' and "
      query << "Units='J';"
      query_results = sqlFile.execAndReturnFirstDouble(query)
      if query_results.empty?
        runner.registerWarning("Did not find value for monthly heating for #{month}.")
        return false
      else
        source_units = 'J'
        target_units = 'kWh'
        value = OpenStudio.convert(query_results.get, source_units, target_units).get
        monthly_htg_cons << value.round(2)
      end

      # gather monthly heating from BUILDING ENERGY PERFORMANCE - DISTRICT COOLING Custom Monthly Report
      query = 'SELECT Value FROM tabulardatawithstrings WHERE '
      query << "ReportName='BUILDING ENERGY PERFORMANCE - DISTRICT COOLING' and "
      query << "ReportForString='Meter' and "
      query << "TableName='Custom Monthly Report' and "
      query << "RowName='#{month}' and "
      query << "ColumnName='COOLING:DISTRICTCOOLING' and "
      query << "Units='J';"
      query_results = sqlFile.execAndReturnFirstDouble(query)
      if query_results.empty?
        runner.registerWarning("Did not find value for monthly cooling for #{month}.")
        return false
      else
        source_units = 'J'
        target_units = 'kWh'
        value = OpenStudio.convert(query_results.get, source_units, target_units).get
        monthly_clg_cons << value.round(2)
      end

      # gather monthly heating from BUILDING ENERGY PERFORMANCE - DISTRICT HEATING PEAK DEMAND Custom Monthly Report
      query = 'SELECT Value FROM tabulardatawithstrings WHERE '
      query << "ReportName='BUILDING ENERGY PERFORMANCE - DISTRICT HEATING WATER PEAK DEMAND' and "
      query << "ReportForString='Meter' and "
      query << "TableName='Custom Monthly Report' and "
      query << "RowName='#{month}' and "
      query << "ColumnName='DISTRICTHEATINGWATER:FACILITY {Maximum}' and "
      query << "Units='W';"
      query_results = sqlFile.execAndReturnFirstDouble(query)
      if query_results.empty?
        runner.registerWarning("Did not find value for monthly heating peak value for #{month}.")
        return false
      else
        source_units = 'W'
        target_units = 'kW'
        value = OpenStudio.convert(query_results.get, source_units, target_units).get
        monthly_htg_peak_val << value.round(2)
      end
      # time of peak
      query = 'SELECT Value FROM tabulardatawithstrings WHERE '
      query << "ReportName='BUILDING ENERGY PERFORMANCE - DISTRICT HEATING WATER PEAK DEMAND' and "
      query << "ReportForString='Meter' and "
      query << "TableName='Custom Monthly Report' and "
      query << "RowName='#{month}' and "
      query << "ColumnName LIKE '%DISTRICTHEATINGWATER:FACILITY {TIMESTAMP}%' and "
      query << "Units='';"
      query_results = sqlFile.execAndReturnFirstString(query)
      if query_results.empty?
        runner.registerWarning("Did not find value for monthly heating peak time for #{month}.")
        return false
      else
        value = query_results.get
        monthly_htg_peak_time << value
      end

      # gather monthly heating from BUILDING ENERGY PERFORMANCE - DISTRICT HEATING PEAK DEMAND Custom Monthly Report
      query = 'SELECT Value FROM tabulardatawithstrings WHERE '
      query << "ReportName='BUILDING ENERGY PERFORMANCE - DISTRICT COOLING PEAK DEMAND' and "
      query << "ReportForString='Meter' and "
      query << "TableName='Custom Monthly Report' and "
      query << "RowName='#{month}' and "
      query << "ColumnName='DISTRICTCOOLING:FACILITY {Maximum}' and "
      query << "Units='W';"
      query_results = sqlFile.execAndReturnFirstDouble(query)
      if query_results.empty?
        runner.registerWarning("Did not find value for monthly cooling peak value for #{month}.")
        return false
      else
        source_units = 'W'
        target_units = 'kW'
        value = OpenStudio.convert(query_results.get, source_units, target_units).get
        monthly_clg_peak_val << value.round(2)
      end
      # time of peak
      query = 'SELECT Value FROM tabulardatawithstrings WHERE '
      query << "ReportName='BUILDING ENERGY PERFORMANCE - DISTRICT COOLING PEAK DEMAND' and "
      query << "ReportForString='Meter' and "
      query << "TableName='Custom Monthly Report' and "
      query << "RowName='#{month}' and "
      query << "ColumnName LIKE '%DISTRICTCOOLING:FACILITY {TIMESTAMP}%' and "
      query << "Units='';"
      query_results = sqlFile.execAndReturnFirstString(query)
      if query_results.empty?
        runner.registerWarning("Did not find value for monthly cooling peak time for #{month}.")
        return false
      else
        value = query_results.get
        monthly_clg_peak_time << value
      end
    end

    # populated runnerRegisterValues for each column. Each is Jan-Dec
    runner.registerInfo("Writing out monthly heating and cooling values")
    runner.registerInfo("help: #{monthly_htg_cons.join(",")}")
    runner.registerInfo("help2: #{monthly_clg_cons.join(",")}")
    runner.registerValue("monthly_htg_cons",monthly_htg_cons.join(",")) #kWh
    runner.registerValue("monthly_clg_cons",monthly_clg_cons.join(",")) #kWh
    runner.registerValue("monthly_htg_peak_val",monthly_htg_peak_val.join(",")) #kW
    runner.registerValue("monthly_htg_peak_time",monthly_htg_peak_time.join(",")) #TIMESTAMP
    runner.registerValue("monthly_clg_peak_val",monthly_clg_peak_val.join(",")) #kW
    runner.registerValue("monthly_clg_peak_time",monthly_clg_peak_time.join(",")) #TIMESTAMP

    return @monthly_htg_clg_section
  end

end
