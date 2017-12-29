# see the URL below for information on how to write OpenStudio measures
# http://nrel.github.io/OpenStudio-user-documentation/reference/measure_writing_guide/

#start the measure
class BestestCeReporting < OpenStudio::Ruleset::ReportingUserScript

  # human readable name
  def name
    return "Bestest Ce Reporting"
  end
  # human readable description
  def description
    return "This doesn't generate a user HTML file with any meaningful content. It is here to create runner.RegisterValue objects that will be post processed downstream."
  end
  # human readable description of modeling approach
  def modeler_description
    return "The CSV project for the analysis will be downloaded from the server and then a script will run to pull data into Local Excel File."
  end
  # define the arguments that the user will input
  def arguments()
    args = OpenStudio::Ruleset::OSArgumentVector.new

    # this measure does not require any user arguments, return an empty list

    return args
  end 
  
  # return a vector of IdfObject's to request EnergyPlus objects needed by the run method
  def energyPlusOutputRequests(runner, user_arguments)
    super(runner, user_arguments)

    # get the last idf (just used for building name)
    workspace = runner.lastEnergyPlusWorkspace
    if workspace.empty?
      runner.registerError('Cannot find last idf file.')
      return false
    end
    workspace = workspace.get
    bldg_name = workspace.getObjectsByType("Building".to_IddObjectType).first.getString(0).get

    result = OpenStudio::IdfObjectVector.new

    # Add output requests (consider adding to case hash instead of adding logic here)
    # this gather any non standard output requests. Analysis of output such as binning temps for FF will occur in reporting measure
    # Table 6-1 describes the specific day of results that will be used for testing
    hourly_variables = []

    # adding collection of ranges and timesteps for different variables
    hourly_var_feb = [] # used for Case 1xx and 2xx

    # variables for all CE cases
    hourly_variables << 'Site Outdoor Air Drybulb Temperature'
    hourly_variables << 'Site Outdoor Air Humidity Ratio'
    hourly_variables << 'Surface Inside face Temperature'
    hourly_variables << 'Surface Outside face Temperature'
    hourly_variables << 'Surface Inside Face Convection Heat Transfer Coefficient'
    hourly_variables << 'Surface Outside Face Convection Heat Transfer Coefficient'
    #hourly_variables << 'Zone Mean Air Temperature'
    #hourly_variables << 'Zone Air Temperature'
    hourly_variables << 'Zone Air Humidity Ratio'
    hourly_variables << 'Zone Air System Sensible Heating Energy'
    hourly_variables << 'Zone Air System Sensible Cooling Energy'
    hourly_variables << 'Zone Total Internal Latent Gain Energy'
    hourly_variables << 'Fan Electric Power'
    hourly_variables << 'Fan Rise in Air Temperature'
    hourly_variables << 'Fan Electric Energy'
    #hourly_variables << 'Cooling Coil Total Cooling Energy'
    #hourly_variables << 'Cooling Coil Sensible Cooling Rate'
    #hourly_variables << 'Cooling Coil Sensible Cooling Energy'
    hourly_variables << 'Cooling Coil Electric Power'
    hourly_variables << 'Cooling Coil Electric Energy'
    hourly_variables << 'Cooling Coil Latent Cooling Rate'
    hourly_variables << 'Cooling Coil Latent Cooling Energy'

    # Unitary outputs in these models vs. Zone Window Air Conditioner outputs in legacy
    hourly_variables << 'Unitary System Part Load Ratio'
    hourly_variables << 'Unitary System Total Cooling Rate'
    hourly_variables << 'Unitary System Sensible Cooling Rate'
    #hourly_variables << 'Unitary System Latent Cooling Rate'
    #hourly_variables << 'Unitary System Total Heating Rate'
    #hourly_variables << 'Unitary System Sensible Heating Rate'
    hourly_variables << 'Unitary System Latent Heating Rate'
    hourly_variables << 'Unitary System Ancillary Electric Power'
    hourly_variables << 'Unitary System Dehumidification Induced Heating Demand Rate'
    hourly_variables << 'Unitary System Fan Part Load Ratio'
    hourly_variables << 'Unitary System Compressor Part Load Ratio'
    hourly_variables << 'Unitary System Frost Control Status'

    # variables for EDB and EWB 'Node 6' is the terminal
    # todo - update reporting for 3B to use this
    hourly_variables << 'System Node Temperature'
    hourly_variables << 'System Node Wetbulb Temperature'

    # variables CE 1x through 2x
    if bldg_name.include? "CE1" or bldg_name.include? "CE2"
      hourly_variables << 'Site Outdoor Air Wetbulb Temperature'
      hourly_variables << 'Site Outdoor Air Dewpoint Temperature'
      hourly_variables << 'Site Outdoor Air Enthalpy'
      hourly_variables << 'Site Outdoor Air Relative Humidity'
      hourly_variables << 'Site Outdoor Air Density'
      hourly_variables << 'Site Outdoor Air Barometric Pressure'
      hourly_variables << 'Site Wind Speed'
      hourly_variables << 'Site Direct Solar Radiation Rate per Area'
      hourly_variables << 'Site Diffuse Solar Radiation Rate per Area'

      # hourly for february

      #6.3.1.1 (a,b,c,d)
      hourly_var_feb << 'Air System Electric Energy' #J
      hourly_var_feb << 'Air System DX Cooling Coil Electric Energy' #J
      hourly_var_feb << 'Air System Fan Electric Energy' #J
      # todo - can I get d directly or does d = a - b - c

      #6.3.1.2 (a,b,c)
      hourly_var_feb << 'Cooling Coil Total Cooling Rate' #W
      hourly_var_feb << 'Cooling Coil Sensible Cooling Rate' #W
      hourly_var_feb << 'Cooling Coil Latent Cooling Rate' #W

      #6.3.1.3 (a,b,c)
      hourly_var_feb << 'Unitary System Total Cooling Rate' #W
      hourly_var_feb << 'Unitary System Sensible Cooling Rate' #W
      hourly_var_feb << 'Unitary System Latent Cooling Rate' #W

      #6.3.1.4 (a,b,c)
      # a: calculated from other variables
      hourly_var_feb << 'Zone Mean Air Temperature'
      hourly_var_feb << 'Zone Mean Air Humidity Ratio'


    elsif bldg_name.include? "CE3"
      hourly_variables << 'System Node Temperature'
      hourly_variables << 'System Node Mass Flow Rate'

      # adding same variables to CE3-5 as CD 1-2, but annualy instead of for February
      hourly_variables << 'Air System Electric Energy' #J
      hourly_variables << 'Air System DX Cooling Coil Electric Energy' #J
      hourly_variables << 'Air System Fan Electric Energy' #J
      # todo - can I get d directly or does d = a - b - c
      hourly_variables << 'Cooling Coil Total Cooling Rate' #W
      hourly_variables << 'Cooling Coil Sensible Cooling Rate' #W
      hourly_variables << 'Cooling Coil Latent Cooling Rate' #W
      hourly_variables << 'Unitary System Total Cooling Rate' #W
      hourly_variables << 'Unitary System Sensible Cooling Rate' #W
      hourly_variables << 'Unitary System Latent Cooling Rate' #W
      hourly_variables << 'Zone Mean Air Temperature'
      hourly_variables << 'Zone Mean Air Humidity Ratio'

      # extra for 3x - 5x
      hourly_variables << 'Zone Air Relative Humidity'
      hourly_variables << 'Site Outdoor Air Drybulb Temperature'
      hourly_variables << 'Site Outdoor Air Humidity Ratio'

    elsif bldg_name.include? "CE4" or bldg_name.include? "CE5"
      hourly_variables << 'System Node Temperature'
      hourly_variables << 'System Node Mass Flow Rate'
      hourly_variables << 'System Node Setpoint Temperature'
      hourly_variables << 'Zone Other Equipment Radiant Heating Energy'
      hourly_variables << 'Zone Other Equipment Convective Heating Energy'
      hourly_variables << 'Zone Other Equipment Latent Gain Energy'
      hourly_variables << 'Zone Other Equipment Lost Heat Energy'
      hourly_variables << 'Zone Other Equipment Total Heating Energy'

      # adding same variables to CE3-5 as CD 1-2, but annualy instead of for February
      hourly_variables << 'Air System Electric Energy' #J
      hourly_variables << 'Air System DX Cooling Coil Electric Energy' #J
      hourly_variables << 'Air System Fan Electric Energy' #J
      # todo - can I get d directly or does d = a - b - c
      hourly_variables << 'Cooling Coil Total Cooling Rate' #W
      hourly_variables << 'Cooling Coil Sensible Cooling Rate' #W
      hourly_variables << 'Cooling Coil Latent Cooling Rate' #W
      hourly_variables << 'Unitary System Total Cooling Rate' #W
      hourly_variables << 'Unitary System Sensible Cooling Rate' #W
      hourly_variables << 'Unitary System Latent Cooling Rate' #W
      hourly_variables << 'Zone Mean Air Temperature'
      hourly_variables << 'Zone Mean Air Humidity Ratio'

      # extra for 3x - 5x
      hourly_variables << 'Zone Air Relative Humidity'
      hourly_variables << 'Site Outdoor Air Drybulb Temperature'
      hourly_variables << 'Site Outdoor Air Humidity Ratio'

    else
      runner.registerWarning("Unexpected Case Number")
    end

    hourly_variables.each do |variable|
      result << OpenStudio::IdfObject.load("Output:Variable,,#{variable},hourly;").get
    end

    hourly_var_feb.each do |variable|
      # note: reporting entire runperiod and will grab feb results in post processing
      result << OpenStudio::IdfObject.load("Output:Variable,,#{variable},hourly;").get
    end

    
    return result
  end

  def outputs
    result = OpenStudio::Measure::OSOutputVector.new
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('clg_energy_consumption_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('clg_energy_consumption_compressor')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('clg_energy_consumption_supply_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('clg_energy_consumption_condenser_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evaporator_coil_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evaporator_coil_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evaporator_coil_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('zone_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('zone_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('zone_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_mean_cop')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_mean_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_mean_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_max_cop')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_max_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_max_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_min_cop')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_min_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('feb_min_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_clg_energy_consumption_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_clg_energy_consumption_compressor')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_clg_energy_consumption_supply_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_clg_energy_consumption_condenser_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_evap_coil_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_evap_coil_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_sum_evap_coil_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_mean_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_mean_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_mean_zone_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_mean_zone_relative_humidity')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_mean_odb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('ann_mean_outdoor_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_clg_consumption_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_clg_consumption_compressor')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_clg_consumption_cond_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_clg_consumption_indoor_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_evap_coil_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_evap_coil_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_sum_evap_coil_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_mean_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_mean_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_mean_zone_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('may_sept_mean_zone_relative_humidity')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('energy_consumption_comp_both_fans_wh')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('energy_consumption_comp_both_fans_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('energy_consumption_comp_both_fans_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_sensible_wh')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_sensible_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_sensible_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_latent_wh')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_latent_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_latent_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_sensible_and_latent_wh')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_sensible_and_latent_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('evap_coil_load_sensible_and_latent_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('weather_odb_c')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('weather_odb_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('weather_odb_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('weather_outdoor_humidity_ratio_c')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('weather_outdoor_humidity_ratio_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('weather_outdoor_humidity_ratio_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('cop_2_max_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('cop_2_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('cop_2_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('cop_2_min_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('cop_2_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('cop_2_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('idb_max_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('idb_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('idb_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('idb_min_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('idb_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('idb_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('hr_max_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('hr_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('hr_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('hr_min_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('hr_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('hr_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('rh_max_relative_humidity')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('rh_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('rh_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('rh_min_relative_humidity')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('rh_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('rh_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_cop_2_max_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_cop_2_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_cop_2_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_cop_2_min_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_cop_2_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_cop_2_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_idb_max_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_idb_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_idb_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_idb_min_idb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_idb_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_idb_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_hr_max_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_hr_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_hr_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_hr_min_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_hr_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_hr_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_rh_max_relative_humidity')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_rh_max_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_rh_max_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_rh_min_relative_humidity')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_rh_min_date')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('apr_dec_rh_min_hr')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_energy_consumpton_compressor')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_energy_consumpton_cond_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_evaporator_coil_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_evaporator_coil_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_evaporator_coil_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_zone_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_odb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_edb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_ewb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0628_hourly_outdoor_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_energy_consumption_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_energy_consumption_compressor')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_energy_consumption_supply_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_energy_consumption_condenser_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_evaporator_coil_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_evaporator_coil_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_evaporator_coil_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_zone_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_odb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0430_day_edb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_energy_consumption_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_energy_consumption_compressor')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_energy_consumption_supply_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_energy_consumption_condenser_fan')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_evaporator_coil_load_total')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_evaporator_coil_load_sensible')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_evaporator_coil_load_latent')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_zone_humidity_ratio')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_cop_2')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_odb')
    result << OpenStudio::Measure::OSOutput.makeDoubleOutput('mmdd_0625_day_edb')

    return result
  end

  # define what happens when the measure is run
  def run(runner, user_arguments)
    super(runner, user_arguments)

    # use the built-in error checking 
    if !runner.validateUserArguments(arguments(), user_arguments)
      return false
    end

    sqlFile = runner.lastEnergyPlusSqlFile
    if sqlFile.empty?
      runner.registerError("Cannot find last sql file.")
      return false
    end
    sqlFile = sqlFile.get

    # get the last idf (just used for building name)
    workspace = runner.lastEnergyPlusWorkspace
    if workspace.empty?
      runner.registerError('Cannot find last idf file.')
      return false
    end
    workspace = workspace.get
    bldg_name = workspace.getObjectsByType("Building".to_IddObjectType).first.getString(0).get


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

    # only try to get the annual timeseries if an annual simulation was run
    if ann_env_pd

      # get desired variable
      key_value =  "Environment"
      time_step = "Hourly" # "Zone Timestep", "Hourly", "HVAC System Timestep"
      variable_name = "Site Outdoor Air Drybulb Temperature"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, time_step, variable_name, key_value) # key value would go at the end if we used it.
      
      if output_timeseries.empty?
        runner.registerWarning("Timeseries not found.")
      else
        runner.registerInfo("Found timeseries.")
      end
    else
      runner.registerWarning("No annual environment period found.")
    end

    # this is in CE and Envelope, move to shared resource
    def process_output_timeseries (sqlFile, runner, ann_env_pd, time_step, variable_name, key_value, days_skip_start = 0, days_skip_end = 0)

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

        start_value = 0 + days_skip_start * 24
        end_value = output_values.size - 1 - days_skip_end * 24

        for i in start_value..end_value

          # using this to get average
          array << output_values[i]
          sum += output_values[i]

          # code for min and max
          if min.nil? || output_values[i] < min
            min = output_values[i]
            min_date_time = output_times[i]
          end
          if max.nil? || output_values[i] > max
            max = output_values[i]
            max_date_time = output_times[i]
          end

        end
        return {:array => array, :sum => sum, :avg => sum/array.size.to_f, :min => min, :max => max, :min_date_time => min_date_time, :max_date_time => max_date_time}
      end

    end

    def date_time_parse(date_time)

      array = []

      month = date_time.date.monthOfYear.value
      day_of_month = date_time.date.dayOfMonth.to_s
      hour = date_time.time.hours

      # map month integer to short name
      case month
        when 1
          month = "Jan"
        when 2
          month = "Feb"
        when 3
          month = "Mar"
        when 4
          month = "Apr"
        when 5
          month = "May"
        when 6
          month = "Jun"
        when 7
          month = "Jul"
        when 8
          month = "Aug"
        when 9
          month = "Sep"
        when 10
          month = "Oct"
        when 11
          month = "Nov"
        when 12
          month = "Dec"
      end

      array << "#{"%02d" % day_of_month}-#{month}"
      array << hour

      return array

    end

    # date format should be dd-MMM. Hour is integer
    # note: This is also used in the bestest_populate_report.rb for envelope
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

    # add runner.registerValues for bestest reporting 5-3A
    if bldg_name.include? "CE1" or bldg_name.include? "CE2"

      # get clg_energy_consumption_total
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      total_cooling_energy_consumption_j = timeseries_hash
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('clg_energy_consumption_total',value_kwh)
      # get clg_energy_consumption_compressor
      variable_name = "Air System DX Cooling Coil Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('clg_energy_consumption_compressor',value_kwh)
      # get clg_energy_consumption_supply_fan
      variable_name = "Air System Fan Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('clg_energy_consumption_supply_fan',value_kwh)
      # todo - can I get d directly or does d = a - b - c
      runner.registerValue('clg_energy_consumption_condenser_fan','tbd')

      # get evaporator_coil_load_total
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
        variable_name = "Cooling Coil Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('evaporator_coil_load_total',value_kwh)
      # get evaporator_coil_load_sensible
      variable_name = "Cooling Coil Sensible Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('evaporator_coil_load_sensible',value_kwh)
      # get evaporator_coil_load_latent
      variable_name = "Cooling Coil Latent Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('evaporator_coil_load_latent',value_kwh)

      # get zone_load_total
      key_value =  "AIR LOOP HVAC UNITARY SYSTEM 1"
      variable_name = "Unitary System Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      net_refrigeration_effect_w = timeseries_hash
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('zone_load_total',value_kwh)
      # get zone_load_sensible
      variable_name = "Unitary System Sensible Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('zone_load_sensible',value_kwh)
      # get zone_load_latent
      variable_name = "Unitary System Latent Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('zone_load_latent',value_kwh)

      # get feb_cop
      mean_cop = net_refrigeration_effect_w[:avg] / (total_cooling_energy_consumption_j[:avg]/3600.0) # W = J/s
      runner.registerValue('feb_mean_cop',mean_cop)
      cop_array = []
      total_cooling_energy_consumption_j[:array].size.times.each do |i|
        cop_array << net_refrigeration_effect_w[:array][i]/(total_cooling_energy_consumption_j[:array][i]/3600.0) # W = J/s
      end
      runner.registerValue('feb_max_cop',cop_array.max)
      runner.registerValue('feb_min_cop',cop_array.min)
      # get feb_idb
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      runner.registerValue('feb_mean_idb',timeseries_hash[:avg])
      runner.registerValue('feb_max_idb',timeseries_hash[:max])
      runner.registerValue('feb_min_idb',timeseries_hash[:min])
      # get feb_humidity_ratio
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,31,306)
      runner.registerValue('feb_mean_humidity_ratio',timeseries_hash[:avg])
      runner.registerValue('feb_max_humidity_ratio',timeseries_hash[:max])
      runner.registerValue('feb_min_humidity_ratio',timeseries_hash[:min])

    else

      # Annual Sums and Means Table

      # get clg_energy_consumption_total
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('ann_sum_clg_energy_consumption_total',value_kwh)
      # get clg_energy_consumption_compressor (this includes the compressor fan)
      variable_name = "Air System DX Cooling Coil Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      # store avg value for cop_2
      compressor_and_outdoor_fan = value_kwh
      runner.registerValue('ann_sum_clg_energy_consumption_compressor',value_kwh)
      # indoor fan
      variable_name = "Air System Fan Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('ann_sum_clg_energy_consumption_supply_fan',value_kwh)
      # get clg_energy_consumption_supply_fan  (can't calculate this)
      runner.registerValue('ann_sum_clg_energy_consumption_condenser_fan','tbd')
      # get evaporator_coil_load_total
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      # store avg value for cop_2
      evaporator_coil_load = value_kwh
      runner.registerValue('ann_sum_evap_coil_load_total',value_kwh)
      # get evaporator_coil_load_sensible
      variable_name = "Cooling Coil Sensible Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('ann_sum_evap_coil_load_sensible',value_kwh)
      # get evaporator_coil_load_latent
      variable_name = "Cooling Coil Latent Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('ann_sum_evap_coil_load_latent',value_kwh)
      # get cop_2
      mean_cop_2 = evaporator_coil_load/compressor_and_outdoor_fan
      runner.registerValue('ann_mean_cop_2',mean_cop_2)
      # get idb
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('ann_mean_idb',timeseries_hash[:avg])
      # get humidity_ratio
      variable_name = "Zone Mean Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('ann_mean_zone_humidity_ratio',timeseries_hash[:avg])
      # get relative_humidity
      variable_name = "Zone Air Relative Humidity"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('ann_mean_zone_relative_humidity',timeseries_hash[:avg])
      # get site avg odb
      key_value =  "Environment"
      variable_name = "Site Outdoor Air Drybulb Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('ann_mean_odb',timeseries_hash[:avg])
      # get site avg humidity ratio
      variable_name = "Site Outdoor Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('ann_mean_outdoor_humidity_ratio',timeseries_hash[:avg])

      # populate may_sept data
      # get clg_energy_consumption_total
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('may_sept_sum_clg_consumption_total',value_kwh)
      # get clg_energy_consumption_compressor (this includes the compressor fan)
      variable_name = "Air System DX Cooling Coil Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      # store avg value for cop_2
      compressor_and_outdoor_fan = value_kwh
      runner.registerValue('may_sept_sum_clg_consumption_compressor',value_kwh)
      variable_name = "Air System Fan Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'J','kWh').get
      runner.registerValue('may_sept_sum_clg_consumption_indoor_fan',value_kwh)
      # get clg_energy_consumption_supply_fan (can't calculate this)
      runner.registerValue('may_sept_sum_clg_consumption_cond_fan','tbd')
      # get evaporator_coil_load_total
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      # store avg value for cop_2
      evaporator_coil_load = value_kwh
      runner.registerValue('may_sept_sum_evap_coil_load_total',value_kwh)
      # get evaporator_coil_load_sensible
      variable_name = "Cooling Coil Sensible Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('may_sept_sum_evap_coil_load_sensible',value_kwh)
      # get evaporator_coil_load_latent
      variable_name = "Cooling Coil Latent Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      value_kwh = OpenStudio.convert(timeseries_hash[:sum],'Wh','kWh').get
      runner.registerValue('may_sept_sum_evap_coil_load_latent',value_kwh)
      # get cop_2
      mean_cop_2 = evaporator_coil_load/compressor_and_outdoor_fan
      runner.registerValue('may_sept_mean_cop_2',mean_cop_2)
      # get idb
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      runner.registerValue('may_sept_mean_idb',timeseries_hash[:avg])
      # get humidity_ratio
      variable_name = "Zone Mean Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      runner.registerValue('may_sept_mean_zone_humidity_ratio',timeseries_hash[:avg])
      # get relative_humidity
      variable_name = "Zone Air Relative Humidity"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      runner.registerValue('may_sept_mean_zone_relative_humidity',timeseries_hash[:avg])
      # get site avg odb
      key_value =  "Environment"
      variable_name = "Site Outdoor Air Drybulb Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      runner.registerValue('may_sept_mean_odb',timeseries_hash[:avg])
      # get site avg humidity ratio
      variable_name = "Site Outdoor Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,120,92)
      runner.registerValue('may_sept_mean_outdoor_humidity_ratio',timeseries_hash[:avg])

      # Annual Hourly Integrated Maxima Consumptions and Loads Table

      # get supply_fan
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_wh = OpenStudio.convert(timeseries_hash[:max],'J','Wh').get
      runner.registerValue('energy_consumption_comp_both_fans_wh',value_wh)
      runner.registerValue('energy_consumption_comp_both_fans_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('energy_consumption_comp_both_fans_hr',date_time_parse(timeseries_hash[:max_date_time])[1])

      # get evaporator_coil
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Sensible Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_wh = OpenStudio.convert(timeseries_hash[:max],'Wh','Wh').get
      runner.registerValue('evap_coil_load_sensible_wh',value_wh)
      runner.registerValue('evap_coil_load_sensible_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('evap_coil_load_sensible_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      variable_name = "Cooling Coil Latent Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_wh = OpenStudio.convert(timeseries_hash[:max],'Wh','Wh').get
      runner.registerValue('evap_coil_load_latent_wh',value_wh)
      runner.registerValue('evap_coil_load_latent_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('evap_coil_load_latent_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      variable_name = "Cooling Coil Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      value_wh = OpenStudio.convert(timeseries_hash[:max],'Wh','Wh').get
      runner.registerValue('evap_coil_load_sensible_and_latent_wh',value_wh)
      runner.registerValue('evap_coil_load_sensible_and_latent_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('evap_coil_load_sensible_and_latent_hr',date_time_parse(timeseries_hash[:max_date_time])[1])

      key_value =  "Environment"
      variable_name = "Site Outdoor Air Drybulb Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('weather_odb_c',timeseries_hash[:max])
      runner.registerValue('weather_odb_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('weather_odb_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      variable_name = "Site Outdoor Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('weather_outdoor_humidity_ratio_c',timeseries_hash[:max])
      runner.registerValue('weather_outdoor_humidity_ratio_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('weather_outdoor_humidity_ratio_hr',date_time_parse(timeseries_hash[:max_date_time])[1])

      # Annual Hourly Integrated Maxima - cop_2 and Zone Table

      # get denominator for cop_2
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System DX Cooling Coil Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      compressor_and_outdoor_fan_array = timeseries_hash[:array]
      # get nominator for cop_2
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      evaporator_coil_load_array = timeseries_hash[:array]
      # calculate 8760 cop
      cop_2_8760 = []
      compressor_and_outdoor_fan_array.size.times.each do |i|
        if compressor_and_outdoor_fan_array[i] > 0.0
          cop_2_8760 << evaporator_coil_load_array[i] / (compressor_and_outdoor_fan_array[i]/3600.0) # W = J/s
        else
          cop_2_8760 << 0.0 # don't like putting value here but if I don't put value can't get min and max, and if I skip entry index position will be wrong
        end
      end
      remove_zeros = {}
      cop_2_8760.each.with_index do |value,i|
        if value > 0
          remove_zeros[i] = value
        end
      end
      index_of_max = cop_2_8760.each_index.max
      hash_min = remove_zeros.values.min
      runner.registerValue('cop_2_max_cop_2',cop_2_8760.max)
      runner.registerValue('cop_2_max_date',return_date_time_from_8760_index(index_of_max)[0])
      runner.registerValue('cop_2_max_hr',return_date_time_from_8760_index(index_of_max)[1])
      runner.registerValue('cop_2_min_cop_2',hash_min)
      runner.registerValue('cop_2_min_date',return_date_time_from_8760_index(remove_zeros.key(hash_min))[0])
      runner.registerValue('cop_2_min_hr',return_date_time_from_8760_index(remove_zeros.key(hash_min))[1])
      # get idb
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('idb_max_idb',timeseries_hash[:max])
      runner.registerValue('idb_max_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('idb_max_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      runner.registerValue('idb_min_idb',timeseries_hash[:min])
      runner.registerValue('idb_min_date',date_time_parse(timeseries_hash[:min_date_time])[0])
      runner.registerValue('idb_min_hr',date_time_parse(timeseries_hash[:min_date_time])[1])
      # get humidity_ratio
      variable_name = "Zone Mean Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('hr_max_humidity_ratio',timeseries_hash[:max])
      runner.registerValue('hr_max_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('hr_max_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      runner.registerValue('hr_min_humidity_ratio',timeseries_hash[:min])
      runner.registerValue('hr_min_date',date_time_parse(timeseries_hash[:min_date_time])[0])
      runner.registerValue('hr_min_hr',date_time_parse(timeseries_hash[:min_date_time])[1])
      # get relative_humidity
      variable_name = "Zone Air Relative Humidity"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('rh_max_relative_humidity',timeseries_hash[:max])
      runner.registerValue('rh_max_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('rh_max_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      runner.registerValue('rh_min_relative_humidity',timeseries_hash[:min])
      runner.registerValue('rh_min_date',date_time_parse(timeseries_hash[:min_date_time])[0])
      runner.registerValue('rh_min_hr',date_time_parse(timeseries_hash[:min_date_time])[1])

      # populate april_dec data
      # get denominator for cop_2
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System DX Cooling Coil Electric Energy"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,90)
      compressor_and_outdoor_fan_array = timeseries_hash[:array]
      # get nominator for cop_2
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,90)
      evaporator_coil_load_array = timeseries_hash[:array]
      # calculate 8760 cop
      cop_2_8760 = []
      compressor_and_outdoor_fan_array.size.times.each do |i|
        if compressor_and_outdoor_fan_array[i] > 0.0
          cop_2_8760 << evaporator_coil_load_array[i] / (compressor_and_outdoor_fan_array[i]/3600.0) # W = J/s
        else
          cop_2_8760 << 0.0 # don't like putting value here but if I don't put value can't get min and max, and if I skip entry index position will be wrong
        end
      end
      remove_zeros = {}
      cop_2_8760.each.with_index do |value,i|
        if value > 0
          remove_zeros[i] = value
        end
      end
      index_of_max = cop_2_8760.each_index.max
      hash_min = remove_zeros.values.min
      runner.registerValue('apr_dec_cop_2_max_cop_2',cop_2_8760.max)
      runner.registerValue('apr_dec_cop_2_max_date',return_date_time_from_8760_index(index_of_max)[0])
      runner.registerValue('apr_dec_cop_2_max_hr',return_date_time_from_8760_index(index_of_max)[1])
      runner.registerValue('apr_dec_cop_2_min_cop_2',hash_min)
      runner.registerValue('apr_dec_cop_2_min_date',return_date_time_from_8760_index(remove_zeros.key(hash_min))[0])
      runner.registerValue('apr_dec_cop_2_min_hr',return_date_time_from_8760_index(remove_zeros.key(hash_min))[1])
      # get idb
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Temperature"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,90)
      runner.registerValue('apr_dec_idb_max_idb',timeseries_hash[:max])
      runner.registerValue('apr_dec_idb_max_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('apr_dec_idb_max_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      runner.registerValue('apr_dec_idb_min_idb',timeseries_hash[:min])
      runner.registerValue('apr_dec_idb_min_date',date_time_parse(timeseries_hash[:min_date_time])[0])
      runner.registerValue('apr_dec_idb_min_hr',date_time_parse(timeseries_hash[:min_date_time])[1])
      # get humidity_ratio
      variable_name = "Zone Mean Air Humidity Ratio"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,90)
      runner.registerValue('apr_dec_hr_max_humidity_ratio',timeseries_hash[:max])
      runner.registerValue('apr_dec_hr_max_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('apr_dec_hr_max_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      runner.registerValue('apr_dec_hr_min_humidity_ratio',timeseries_hash[:min])
      runner.registerValue('apr_dec_hr_min_date',date_time_parse(timeseries_hash[:min_date_time])[0])
      runner.registerValue('apr_dec_hr_min_hr',date_time_parse(timeseries_hash[:min_date_time])[1])
      # get relative_humidity
      variable_name = "Zone Air Relative Humidity"
      timeseries_hash = process_output_timeseries(sqlFile, runner, ann_env_pd, 'Hourly', variable_name, key_value,90)
      runner.registerValue('apr_dec_rh_max_relative_humidity',timeseries_hash[:max])
      runner.registerValue('apr_dec_rh_max_date',date_time_parse(timeseries_hash[:max_date_time])[0])
      runner.registerValue('apr_dec_rh_max_hr',date_time_parse(timeseries_hash[:max_date_time])[1])
      runner.registerValue('apr_dec_rh_min_relative_humidity',timeseries_hash[:min])
      runner.registerValue('apr_dec_rh_min_date',date_time_parse(timeseries_hash[:min_date_time])[0])
      runner.registerValue('apr_dec_rh_min_hr',date_time_parse(timeseries_hash[:min_date_time])[1])

      # loop to gather hourly data as string from 0628
      def hourly_values(output_timeseries,target_date,source_unit = '',target_unit = '')

        hourly_single_day_array = []
        24.times.each do |i|
          date_string = "#{target_date} #{i+1}:00:00.000"
          date_time = OpenStudio::DateTime.new(date_string)
          val_at_date_time = output_timeseries.get.value(date_time)
          value = OpenStudio.convert(val_at_date_time, source_unit, target_unit).get
          hourly_single_day_array << value

        end
        hourly_single_day_array = hourly_single_day_array.join(',')

        return hourly_single_day_array

      end

      # loop to gather hourly data as string from 0628
      def avg_from_hourly_values(output_timeseries,target_date,source_unit = '',target_unit = '')

        array = hourly_values(output_timeseries,target_date,source_unit,target_unit).split(',')
        array.map! {|item| item.to_f}
        avg = array.reduce(0, :+)/array.size

        return avg

      end

      # Case 300 June 28th Hourly Table

      # get clg_energy_consumption_compressor
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System DX Cooling Coil Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_energy_consumpton_compressor',hourly_values(output_timeseries,'2009-06-28','J','Wh'))
      # get clg_energy_consumption_supply_fan (can't calculate)
      runner.registerValue('mmdd_0628_hourly_energy_consumpton_cond_fan','tbd')
      # get evaporator_coil_load_total
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_evaporator_coil_load_total',hourly_values(output_timeseries,'2009-06-28'))
      # get evaporator_coil_load_sensible
      variable_name = "Cooling Coil Sensible Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_evaporator_coil_load_sensible',hourly_values(output_timeseries,'2009-06-28'))
      # get evaporator_coil_load_latent
      variable_name = "Cooling Coil Latent Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_evaporator_coil_load_latent',hourly_values(output_timeseries,'2009-06-28'))
      # get humidity_ratio
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Humidity Ratio"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_zone_humidity_ratio',hourly_values(output_timeseries,'2009-06-28'))

      # get clg_energy_consumption_total (for cop calc)
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System DX Cooling Coil Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      compressor_and_outdoor_fan_array = hourly_values(output_timeseries,'2009-06-28').split(",")
      # get zone_load_total (for net_refrigeration_effect_w)
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      evaporator_coil_load_array = hourly_values(output_timeseries,'2009-06-28').split(",")
      # calculate cop_2_24
      cop_2_24 = []
      compressor_and_outdoor_fan_array.size.times.each do |i|
        cop_2_24 << evaporator_coil_load_array[i].to_f / (compressor_and_outdoor_fan_array[i].to_f/3600.0) # W = J/s
      end
      runner.registerValue('mmdd_0628_hourly_cop_2',cop_2_24.join(","))

      # get site odb
      key_value =  "Environment"
      variable_name = "Site Outdoor Air Drybulb Temperature"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_odb',hourly_values(output_timeseries,'2009-06-28'))

      # todo - in future navigate air loop to find right node
      if bldg_name.include? "CE300"
        # get return air drybulb and wetbulb
        key_value =  "NODE 9"
        variable_name = "System Node Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        runner.registerValue('mmdd_0628_hourly_edb',hourly_values(output_timeseries,'2009-06-28'))
        key_value =  "NODE 9"
        variable_name = "System Node Wetbulb Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        runner.registerValue('mmdd_0628_hourly_ewb',hourly_values(output_timeseries,'2009-06-28'))
      elsif bldg_name.include? "CE500" or bldg_name.include? "CE530"
        key_value =  "NODE 2"
        variable_name = "System Node Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        runner.registerValue('mmdd_0628_hourly_edb',hourly_values(output_timeseries,'2009-06-28'))
        key_value =  "NODE 2"
        variable_name = "System Node Wetbulb Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        runner.registerValue('mmdd_0628_hourly_ewb',hourly_values(output_timeseries,'2009-06-28'))
      end

      # get site avg humidity ratio
      key_value =  "Environment"
      variable_name = "Site Outdoor Air Humidity Ratio"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      runner.registerValue('mmdd_0628_hourly_outdoor_humidity_ratio',hourly_values(output_timeseries,'2009-06-28'))

      # Case 500 and 530 Average Daily Outputs

      # get clg_energy_consumption_total
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30','J', 'Wh')
      runner.registerValue('mmdd_0430_day_energy_consumption_total',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25','J', 'Wh')
      runner.registerValue('mmdd_0625_day_energy_consumption_total',avg)
      # get clg_energy_consumption_compressor
      variable_name = "Air System DX Cooling Coil Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30','J', 'Wh')
      runner.registerValue('mmdd_0430_day_energy_consumption_compressor',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25','J', 'Wh')
      runner.registerValue('mmdd_0625_day_energy_consumption_compressor',avg)
      # get clg_energy_consumption_supply_fan
      variable_name = "Air System Fan Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30','J', 'Wh')
      runner.registerValue('mmdd_0430_day_energy_consumption_supply_fan',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25','J', 'Wh')
      runner.registerValue('mmdd_0625_day_energy_consumption_supply_fan',avg)

      runner.registerValue('mmdd_0430_day_energy_consumption_condenser_fan','tbd')
      runner.registerValue('mmdd_0625_day_energy_consumption_condenser_fan','tbd')

      # get evaporator_coil_load_total
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
      runner.registerValue('mmdd_0430_day_evaporator_coil_load_total',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
      runner.registerValue('mmdd_0625_day_evaporator_coil_load_total',avg)
      # get evaporator_coil_load_sensible
      variable_name = "Cooling Coil Sensible Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
      runner.registerValue('mmdd_0430_day_evaporator_coil_load_sensible',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
      runner.registerValue('mmdd_0625_day_evaporator_coil_load_sensible',avg)
      # get evaporator_coil_load_latent
      variable_name = "Cooling Coil Latent Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
      runner.registerValue('mmdd_0430_day_evaporator_coil_load_latent',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
      runner.registerValue('mmdd_0625_day_evaporator_coil_load_latent',avg)

      # get humidity_ratio
      key_value =  "ZONE ONE"
      variable_name = "Zone Mean Air Humidity Ratio"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
      runner.registerValue('mmdd_0430_day_zone_humidity_ratio',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
      runner.registerValue('mmdd_0625_day_zone_humidity_ratio',avg)

      # get clg_energy_consumption_total (for cop calc)
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System DX Cooling Coil Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      compressor_and_outdoor_fan_array = hourly_values(output_timeseries,'2009-04-30').split(",")
      compressor_and_outdoor_fan = 0
      compressor_and_outdoor_fan_array.each do |value|
        compressor_and_outdoor_fan += value.to_f
      end
      # get zone_load_total (for net_refrigeration_effect_w)
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      evaporator_coil_load_array = hourly_values(output_timeseries,'2009-04-30').split(",")
      evaporator_coil_load = 0
      evaporator_coil_load_array.each do |value|
        evaporator_coil_load += value.to_f
      end
      # calculate cop_2_24
      cop_2_24 = evaporator_coil_load / (compressor_and_outdoor_fan/3600.0) # W = J/s
      runner.registerValue('mmdd_0430_day_cop_2',cop_2_24)

      # get clg_energy_consumption_total (for cop calc)
      key_value =  "BESTEST CE AIR LOOP"
      variable_name = "Air System DX Cooling Coil Electric Energy"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      compressor_and_outdoor_fan_array = hourly_values(output_timeseries,'2009-06-25').split(",")
      # get zone_load_total (for net_refrigeration_effect_w)
      key_value =  "COIL COOLING DX SINGLE SPEED 1"
      variable_name = "Cooling Coil Total Cooling Rate"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      net_refrigeration_effect_w_array = hourly_values(output_timeseries,'2009-06-25').split(",")
      # calculate cop_2_24
      cop_2_24 = []
      compressor_and_outdoor_fan_array.size.times.each do |i|
        cop_2_24 << evaporator_coil_load_array[i].to_f / (compressor_and_outdoor_fan_array[i].to_f/3600.0) # W = J/s
      end
      runner.registerValue('mmdd_0625_day_cop_2',cop_2_24.reduce(0, :+)/cop_2_24.size)

      # get site avg odb
      key_value =  "Environment"
      variable_name = "Site Outdoor Air Drybulb Temperature"
      output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
      avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
      runner.registerValue('mmdd_0430_day_odb',avg)
      avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
      runner.registerValue('mmdd_0625_day_odb',avg)

      # todo - in future navigate air loop to find right node
      if bldg_name.include? "CE300"
        key_value =  "NODE 9"
        variable_name = "System Node Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
        runner.registerValue('mmdd_0430_day_edb',avg)
        key_value =  "NODE 9"
        variable_name = "System Node Wetbulb Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
        runner.registerValue('mmdd_0625_day_edb',avg)
      elsif bldg_name.include? "CE500" or bldg_name.include? "CE530"
        key_value =  "NODE 2"
        variable_name = "System Node Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        avg = avg_from_hourly_values(output_timeseries,'2009-04-30')
        runner.registerValue('mmdd_0430_day_edb',avg)
        key_value =  "NODE 2"
        variable_name = "System Node Wetbulb Temperature"
        output_timeseries = sqlFile.timeSeries(ann_env_pd, 'Hourly', variable_name, key_value)
        avg = avg_from_hourly_values(output_timeseries,'2009-06-25')
        runner.registerValue('mmdd_0625_day_edb',avg)
      end

    end

    # close the sql file
    sqlFile.close()

    return true
 
  end

end

# register the measure to be used by the application
BestestCeReporting.new.registerWithApplication
