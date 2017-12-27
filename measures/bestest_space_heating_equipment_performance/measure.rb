# see the URL below for information on how to write OpenStudio measures
# http://nrel.github.io/OpenStudio-user-documentation/reference/measure_writing_guide/

# load library to map case to model variables
require "#{File.dirname(__FILE__)}/resources/besttest_case_var_lib"
require "#{File.dirname(__FILE__)}/resources/besttest_model_methods"
require "#{File.dirname(__FILE__)}/resources/epw"

# start the measure
class BestestSpaceHeatingEquipmentPerformance < OpenStudio::Ruleset::ModelUserScript

  # human readable name
  def name
    return "Bestest Space Heating Equipment Performance"
  end
  # human readable description
  def description
    return "Creates test cases described in ASHRAE Standard 140-2014 sections 5.4.1, 5.4.2, and 5.4.3"
  end
  # human readable description of modeling approach
  def modeler_description
    return "This is intended to run on an empty model. It will create the proper model associate it with the proper weather file, and add in necessary output requests. Internally to the measure the test case argument will be mapped to the proper inputs needed to assemble the model. The measure will make some objects on the fly, other objects will be pulled from existing data resources. This measure creates cases described all of section 5.3."
  end
  # define the arguments that the user will input
  def arguments(model)
    args = OpenStudio::Ruleset::OSArgumentVector.new

    #make choice argument for test case
    choices = OpenStudio::StringVector.new
    variable_hash_lookup = BestestCaseVarLib.bestest_5_4_case_defs
    variable_hash_lookup.each do |k,v|
      choices << k
    end

    # creates arg names for spreadsheet
    array = []
    choices.each do |choice|
      array << "'#{choice}'"
    end
    #puts "String for spreadsheet"
    #puts "[#{array.join(",")}]"

    case_num = OpenStudio::Ruleset::OSArgument::makeChoiceArgument("case_num", choices,true)
    case_num.setDisplayName("Test Case Number")
    case_num.setDescription("Measure will generate selected test case.")
    case_num.setDefaultValue('HE100 - Base-case Building and Mechanical Systems')
    args << case_num

    return args
  end

  # define what happens when the measure is run
  def run(model, runner, user_arguments)
    super(model, runner, user_arguments)

    # use the built-in error checking
    if !runner.validateUserArguments(arguments(model), user_arguments)
      return false
    end

    # assign the user inputs to variables
    case_num = runner.getStringArgumentValue("case_num", user_arguments)
    runner.registerInfo("Full case number: #{case_num}")

    # report initial condition of model
    runner.registerInitialCondition("The initial model named #{model.getBuilding.name} has #{model.numObjects} objects.")

    # map case number to arguments and report back arguments
    variable_hash = BestestCaseVarLib.bestest_5_4_case_lookup(case_num,runner)
    if variable_hash == false
      runner.registerError("Didn't find #{case_num} in model variable hash.")
      return false
    else
      # should return one item, get the hash
      variable_hash = variable_hash.first
    end

    # Adjust simulation settings if necessary
    # todo - do I want simple or should this be skipped
    BestestModelMethods.config_sim_settings(runner,model,'Simple','Simple')

    # Add weather file(won't work in apply measures now)
    top_dir = File.dirname(__FILE__)
    weather_dir = "#{top_dir}/resources/"
    weather_file_name = "#{variable_hash[:epw]}WY2.epw"
    weather_file = File.join(weather_dir, weather_file_name)
    epw_file = OpenStudio::EpwFile.new(weather_file)
    weather_object = OpenStudio::Model::WeatherFile.setWeatherFile(model, epw_file).get
    weather_name = "#{epw_file.city}_#{epw_file.stateProvinceRegion}_#{epw_file.country}"
    weather_lat = epw_file.latitude
    weather_lon = epw_file.longitude
    weather_time = epw_file.timeZone
    weather_elev = epw_file.elevation
    site = model.getSite
    site.setName(weather_name)
    site.setLatitude(weather_lat)
    site.setLongitude(weather_lon)
    site.setTimeZone(weather_time)
    site.setElevation(weather_elev)
    runner.registerInfo("Weather > setting weather to #{weather_object.url.get}")

    # need design days for OpenStudio to run, but values should not matter
    summer_design_day = OpenStudio::Model::DesignDay.new(model)
    winter_design_day = OpenStudio::Model::DesignDay.new(model)
    winter_design_day.setDayType('WinterDesignDay')

    # set runperiod
    run_period = model.getRunPeriod
    run_period.setEndMonth(3)
    run_period.setEndDayOfMonth (31)
    runner.registerInfo("Run Period > Setting Simulation Run Period from 1/1 through 3/31.")

    # Lookup envelope
    file_to_clone = nil
    if case_num.include? 'HE'
      file_to_clone = 'Bestest_Geo_HE100.osm'
    else
      runner.registerError("Unexpected Geometry Variables.")
      return false
    end

    # Add envelope from external file
    runner.registerInfo("Envelope > Adding spaces and zones from #{file_to_clone}")
    translator = OpenStudio::OSVersion::VersionTranslator.new
    geo_path = OpenStudio::Path.new(File.dirname(__FILE__) + "/resources/" + "#{file_to_clone}")
    geo_model = translator.loadModel(geo_path).get
    geo_model.getBuilding.clone(model)

    # change heat balance defaults
    model.getHeatBalanceAlgorithm.setMinimumSurfaceConvectionHeatTransferCoefficientValue(0.0000001)

    # Specific to HE cases, set SurfacePropertyConvectionCoefficients
    conv_coef = OpenStudio::Model::SurfacePropertyConvectionCoefficientsMultipleSurface.new(model)
    conv_coef.setSurfaceType('AllExteriorWalls')
    conv_coef.setConvectionCoefficient1Location('Inside')
    conv_coef.setConvectionCoefficient1Type('Value')
    conv_coef.setConvectionCoefficient1(0.1)
    conv_coef.setConvectionCoefficient2Location('Outside')
    conv_coef.setConvectionCoefficient2Type('Value')
    conv_coef.setConvectionCoefficient2(0.0000001)

    conv_coef = OpenStudio::Model::SurfacePropertyConvectionCoefficientsMultipleSurface.new(model)
    conv_coef.setSurfaceType('AllExteriorRoofs')
    conv_coef.setConvectionCoefficient1Location('Inside')
    conv_coef.setConvectionCoefficient1Type('Value')
    conv_coef.setConvectionCoefficient1(20.0)
    conv_coef.setConvectionCoefficient2Location('Outside')
    conv_coef.setConvectionCoefficient2Type('Value')
    conv_coef.setConvectionCoefficient2(20.0)

    # Load resource file
    file_resource = "bestest_resources.osm"
    runner.registerInfo("Shared Resources > Loading #{file_resource}")
    translator = OpenStudio::OSVersion::VersionTranslator.new
    resource_path = OpenStudio::Path.new(File.dirname(__FILE__) + "/resources/" + "#{file_resource}")
    resource_model = translator.loadModel(resource_path).get

    # no internal loads in HE cases

    # no infiltration in HE cases

    # setup clg thermostat schedule
    bestest_no_clg = resource_model.getModelObjectByName("No Cooling").get.to_ScheduleRuleset.get
    clg_setp = bestest_no_clg.clone(model).to_ScheduleRuleset.get

    # setup htg thermostat schedule
    if variable_hash[:htg_set].is_a? Float
      htg_setp = OpenStudio::Model::ScheduleConstant.new(model)
      htg_setp.setValue(variable_hash[:htg_set])
      htg_setp.setName("#{variable_hash[:htg_set]} C")
    elsif variable_hash[:htg_set] == [15.0,20.0] # HE220 and HE230 use same htg setpoint schedule
      resource_sch = resource_model.getModelObjectByName("HE220_htg").get.to_ScheduleRuleset.get
      htg_setp = resource_sch.clone(model).to_ScheduleRuleset.get
    else
      runner.registerError("Unexpected heating setpoint variable")
      return false
    end

    # create thermostats
    thermostat = OpenStudio::Model::ThermostatSetpointDualSetpoint.new(model)
    thermostat.setCoolingSetpointTemperatureSchedule(clg_setp)
    thermostat.setHeatingSetpointTemperatureSchedule(htg_setp)
    zone = model.getThermalZones.first
    zone.setThermostatSetpointDualSetpoint(thermostat)
    runner.registerInfo("Thermostat > #{zone.name} has clg setpoint sch named #{clg_setp.name} and htg setpoint sch named #{htg_setp.name}.")

    # add in HVAC
    BestestModelMethods.create_he_system(runner,model,variable_hash)

    # rename the building
    model.getBuilding.setName("BESTEST Case #{case_num}")
    runner.registerInfo("Renaming Building > #{model.getBuilding.name}")

    # set timesteps per hour
    timestep = model.getTimestep
    timestep.setNumberOfTimestepsPerHour(4)

    # note: set interior solar distribution fractions isn't needed if E+ auto calculates it

    # note: moved variable requests to reporting measure in energyPlusOutputReqests method

    # report final condition of model
    runner.registerFinalCondition("The final model named #{model.getBuilding.name} has #{model.numObjects} objects.")

    return true

  end
  
end

# register the measure to be used by the application
BestestSpaceHeatingEquipmentPerformance.new.registerWithApplication
