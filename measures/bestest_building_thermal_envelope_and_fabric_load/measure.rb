# see the URL below for information on how to write OpenStudio measures
# http://nrel.github.io/OpenStudio-user-documentation/reference/measure_writing_guide/

# load library to map case to model variables
require "#{File.dirname(__FILE__)}/resources/besttest_case_var_lib"
require "#{File.dirname(__FILE__)}/resources/besttest_model_methods"
require "#{File.dirname(__FILE__)}/resources/epw"

# start the measure
class BestestBuildingThermalEnvelopeAndFabricLoad < OpenStudio::Ruleset::ModelUserScript

  # human readable name
  def name
    return "Bestest Building Thermal Envelope and Fabric Load"
  end

  # human readable description
  def description
    return "Creates test cases described in ASHRAE Standard 140-2014 sections 5.2.1, 5.2.2, and 5.2.3."
  end

  # human readable description of modeling approach
  def modeler_description
    return "This is intended to run on an empty model. It will create the proper model associate it with the proper weather file, and add in necessary output requests. Internally to the measure the test case argument will be mapped to the proper inputs needed to assemble the model. The measure will make some objects on the fly, other objects will be pulled from existing data resources. This measure creates cases described all of section 5.2 except for section 5.2.4 - Ground-Coupled Slab-on-Grade Analytical Verification Tests."
  end

  # define the arguments that the user will input
  def arguments(model)
    args = OpenStudio::Ruleset::OSArgumentVector.new

    #make choice argument for test case
    choices = OpenStudio::StringVector.new
    variable_hash_lookup = BestestCaseVarLib.bestest_5_2_3_case_defs
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
    case_num.setDefaultValue("600 - Base Case")
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
    variable_hash = BestestCaseVarLib.bestest_5_2_3_case_lookup(case_num,runner)
    if variable_hash == false
      runner.registerError("Didn't find #{case_num} in model variable hash.")
      return false
    else
      # should return one item, get the hash
      variable_hash = variable_hash.first
    end

    # Adjust simulation settings if necessary
    BestestModelMethods.config_sim_settings(runner,model,'TARP','DOE-2')

    # Add weather file and design day objects (won't work in apply measures now)
    top_dir = File.dirname(__FILE__)
    weather_dir = "#{top_dir}/resources/"
    weather_file_name = "DRYCOLDTMY.epw"
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

    # Lookup envelope
    file_to_clone = nil
    if variable_hash[:custom] == true and case_num.include? '960'
      # add in sun space geometry
      file_to_clone = 'Bestest_Geo_Sunspace.osm'
    elsif case_num.include? '195' or case_num.include? '395'
      # add in geometry with no fenestration
      file_to_clone = 'Bestest_Geo_South_0_0_0.osm'
    elsif variable_hash[:orient] == 'S'
      if variable_hash[:shade] == false
        # add in south glazing without an overhang
        file_to_clone = 'Bestest_Geo_South_12_0_0.osm'
      elsif variable_hash[:shade] == 1.0 and variable_hash[:shade_type] == 'H'
        # add in south glazing with an overhang
        file_to_clone = 'Bestest_Geo_South_12_1_0.osm'
      else
        runner.registerError("Unexpected Geometry Variables for South Overhangs.")
        return false
      end
    elsif variable_hash[:orient] == 'E,W'
      if variable_hash[:shade] == false
        # add in east/west glazing without an overhang
        file_to_clone = 'Bestest_Geo_EastWest_6_0_0.osm'
      elsif variable_hash[:shade] == 1.0 and variable_hash[:shade_type] == 'H,V'
        # add in east/west glazing with an overhang
        file_to_clone = 'Bestest_Geo_EastWest_6_1_1.osm'
      else
        runner.registerError("Unexpected Geometry Variables for East/West Overhangs.")
        return false
      end
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

    if case_num.include? ('960')

      name_a = 'SUN ZONE SURFACE NORTH'
      name_b = 'ZONE SURFACE SOUTH'
      surface_a = nil
      surface_b = nil

      model.getSurfaces.each do |surface|
        puts surface.name
        if surface.name.to_s == name_a
          surface_a = surface
        end
        if surface.name.to_s == name_b
          surface_b = surface
        end
      end

      if surface_a.nil? || surface_b.nil?
        runner.registerError("Didn't find expected surfaces")
      else
        surface_a.setAdjacentSurface (surface_b)
      end
    end

    # if no windows then replace winodws with doors except for case 195 and 395
    if variable_hash[:glass_area].nil? || variable_hash[:glass_area] == 0.0
      no_windows = true
    else
      no_windows = false
    end
    if case_num.include?('195') || case_num.include?('395') || case_num.include?('960')
      add_doors = false
    else
      add_doors = true
    end
    if no_windows and add_doors
      model.getSubSurfaces.each do |sub_surface|
        sub_surface.setSubSurfaceType("Door")
      end
      runner.registerInfo("Envelope > Replacing windows with doors as high conductance walls.")
    end

    # Load resource file
    file_resource = "bestest_resources.osm"
    runner.registerInfo("Shared Resources > Loading #{file_resource}")
    translator = OpenStudio::OSVersion::VersionTranslator.new
    resource_path = OpenStudio::Path.new(File.dirname(__FILE__) + "/resources/" + "#{file_resource}")
    resource_model = translator.loadModel(resource_path).get

    # Lookup construction sets
    const_sets_to_clone = [] # for 960 two const. sets to use
    if variable_hash[:mass] == "L"
      const_sets_to_clone << "BESTEST LT"
    elsif variable_hash[:mass] == "H"
      const_sets_to_clone << "BESTEST HW"
    elsif variable_hash[:custom] == true and case_num.include? '960'
      const_sets_to_clone << "BESTEST LT"
      const_sets_to_clone << "BESTEST HW"
    else
      runner.registerError("Unexpected mass value.")
      return false
    end

    # Add construction sets
    const_sets = []
    const_sets_to_clone.each do |construction_set_name|

      # this method doesn't work on server 1.10.6 but is more direct
      #const_set = resource_model.getModelObjectByName(construction_set_name).get.to_DefaultConstructionSet.get

      resource_model.getDefaultConstructionSets.each do |res_const_set|
        if construction_set_name == res_const_set.name.to_s
          const_set = res_const_set.clone(model).to_DefaultConstructionSet.get
          const_sets << const_set
        end
      end
    end
    if const_sets.size == 1
      model.getBuilding.setDefaultConstructionSet(const_sets.first)
      runner.registerInfo("Constructions > Setting #{const_sets.first.name} as the default construction set for the building.")
    else
      model.getDefaultConstructionSets.each do |const_set|
        if const_set.name.to_s == "BESTEST LT"
          model.getBuilding.setDefaultConstructionSet(const_set)
          runner.registerInfo("Constructions > Setting #{const_set.name} as the default construction set for the building.")
        elsif const_set.name.to_s == "BESTEST HW"
          # set default construction set for sun space
          model.getSpaces.each do |space|
            if space.name.to_s == "SUN ZONE Space"
              space.setDefaultConstructionSet(const_set)
              runner.registerInfo("Constructions > Setting #{const_set.name} as the construction set for #{space.name}.")
            end
          end
        end
      end
    end

    # set opaque surface properties (no special logic needed for sun space. Internal wall from const set is correct)
    altered_materials =  BestestModelMethods.set_opqaue_surface_properties(model,variable_hash)
    runner.registerInfo("Surface Properties > altered #{altered_materials.uniq.size} materials.")

    # lookup schedules that might be needed
    bestest_htg_setback = resource_model.getModelObjectByName("BESTEST htg SETBACK").get.to_ScheduleRuleset.get
    bestest_clg_night_vent = resource_model.getModelObjectByName("BESTEST clg Night Vent").get.to_ScheduleRuleset.get
    bestest_night_vent = resource_model.getModelObjectByName("BESTEST Night Vent").get.to_ScheduleRuleset.get
    bestest_no_htg = resource_model.getModelObjectByName("No Heating").get.to_ScheduleRuleset.get
    bestest_no_clg = resource_model.getModelObjectByName("No Cooling").get.to_ScheduleRuleset.get
    always_on = resource_model.getModelObjectByName("Always On Ruleset Fractional").get.to_ScheduleRuleset.get

    # Add internal loads
    if variable_hash[:int_gen] and variable_hash[:int_gen] > 0.0
      # clone def, create instance and ad to spae
      res_cother_equip = resource_model.getModelObjectByName("ZONE ONE OthEq 1 Definition").get.to_OtherEquipmentDefinition.get
      other_equip_def = res_cother_equip.clone(model).to_OtherEquipmentDefinition.get
      load_inst = OpenStudio::Model::OtherEquipment.new(other_equip_def)
      load_inst.setSchedule(always_on.clone(model).to_ScheduleRuleset.get)
      load_inst.setSpace(model.getSpaces.first)
      runner.registerInfo("Internal Loads > Adding #{other_equip_def.name} to #{model.getSpaces.first.name}.")
    elsif variable_hash[:custom]
      res_cother_equip = resource_model.getModelObjectByName("ZONE ONE OthEq 1 Definition").get.to_OtherEquipmentDefinition.get
      other_equip_def = res_cother_equip.clone(model).to_OtherEquipmentDefinition.get
      load_inst = OpenStudio::Model::OtherEquipment.new(other_equip_def)
      load_inst.setSchedule(always_on.clone(model).to_ScheduleRuleset.get)
      space = model.getModelObjectByName("BACK ZONE Space").get.to_Space.get
      load_inst.setSpace(space)
      runner.registerInfo("Internal Loads > Adding #{other_equip_def.name} to #{space.name}.")
    else
      runner.registerInfo("Internal Loads > No Other Eqipment Loads added")
    end

    # Add infiltration
    if variable_hash[:infil]
      ach = variable_hash[:infil]
    else
      ach = 0.5 # confirm this is what sunspace needs
    end
    model.getSpaces.each do |space|
      infil = OpenStudio::Model::SpaceInfiltrationDesignFlowRate.new(model)
      infil.setSpace(space)

      # should be true for 650,950,650FF,950FF
      if variable_hash[:vent]
        infil.setDesignFlowRate (0.4911) # value from legacy IDF. Combined infiltraiton and night ventilation
        infil.setSchedule(bestest_night_vent.clone(model).to_ScheduleRuleset.get)
      else
        infil.setAirChangesperHour(ach)
        infil.setSchedule(always_on.clone(model).to_ScheduleRuleset.get)
      end

      runner.registerInfo("Infiltration > Setting to #{infil.airChangesperHour} ACH for #{space.name}.")
    end

    # setup clg thermostat schedule
    clg_setp = nil
    if variable_hash[:vent] and variable_hash[:clg_set].is_a? Float
      clg_setp = bestest_clg_night_vent.clone(model).to_ScheduleRuleset.get
    elsif variable_hash[:clg_set].is_a? Float
      clg_setp = OpenStudio::Model::ScheduleConstant.new(model)
      clg_setp.setValue(variable_hash[:clg_set])
      clg_setp.setName("#{variable_hash[:clg_set]} C")
    elsif variable_hash[:clg_set] == "SETBACK"
      clg_setp = OpenStudio::Model::ScheduleConstant.new(model)
      clg_setp.setValue(27.0)
      clg_setp.setName("27.0 C")
    elsif variable_hash[:clg_set] == "NONE" and variable_hash[:htg_set] == "NONE"
      # leave as nil, thermostat won't be made
    elsif variable_hash[:clg_set] == "NONE"
      clg_setp = bestest_no_clg.clone(model).to_ScheduleRuleset.get
    elsif variable_hash[:custom]
      clg_setp = OpenStudio::Model::ScheduleConstant.new(model)
      clg_setp.setValue(27.0)
      clg_setp.setName("27.0 C")
    else
      runner.registerError("Unexpected cooling setpoint value.")
      return false
    end

    # setup htg thermostat schedule
    htg_setp = nil
    if variable_hash[:htg_set].is_a? Float
      htg_setp = OpenStudio::Model::ScheduleConstant.new(model)
      htg_setp.setValue(variable_hash[:htg_set])
      htg_setp.setName("#{variable_hash[:htg_set]} C")
    elsif variable_hash[:htg_set] == "SETBACK"
      htg_setp = bestest_htg_setback.clone(model).to_ScheduleRuleset.get
    elsif variable_hash[:htg_set] == "NONE" and variable_hash[:clg_set] == "NONE"
      # leave as nil, thermostat won't be made
    elsif variable_hash[:htg_set] == "NONE"
      htg_setp = bestest_no_htg.clone(model).to_ScheduleRuleset.get
    elsif variable_hash[:custom]
      htg_setp = OpenStudio::Model::ScheduleConstant.new(model)
      htg_setp.setValue(20.0)
      htg_setp.setName("20.0 C")
    else
      runner.registerError("Unexpected heating setpoint value.")
      return false
    end

    # create thermostats
    model.getThermalZones.each do |zone|
      next if clg_setp.nil? || htg_setp.nil?
      next if zone.name.to_s == "SUN ZONE"
      thermostat = OpenStudio::Model::ThermostatSetpointDualSetpoint.new(model)
      thermostat.setCoolingSetpointTemperatureSchedule(clg_setp)
      thermostat.setHeatingSetpointTemperatureSchedule(htg_setp)
      zone.setThermostatSetpointDualSetpoint(thermostat)
      runner.registerInfo("Thermostat > #{zone.name} has clg setpoint sch named #{clg_setp.name} and htg setpoint sch named #{htg_setp.name}.")
    end

    # add in night ventilation (added as extra infiltration)

    # add in HVAC
    if !variable_hash[:ff]
      model.getThermalZones.each do |zone|
        next if zone.name.to_s == "SUN ZONE"

        # doesn't give control I need
        #zone.setUseIdealAirLoads(true)

        # use this instead
        ideal_loads = OpenStudio::Model::ZoneHVACIdealLoadsAirSystem.new(model)
        ideal_loads.addToThermalZone(zone)
        ideal_loads.setMaximumHeatingSupplyAirHumidityRatio(0.01)
        ideal_loads.setMinimumCoolingSupplyAirHumidityRatio(0.01)
        ideal_loads.setDehumidificationControlType('ConstantSupplyHumidityRatio')
        ideal_loads.setHumidificationControlType('ConstantSupplyHumidityRatio')

        runner.registerInfo("HVAC > Adding ideal air loads to #{zone.name}.")
      end
    end

    # rename the building
    model.getBuilding.setName("BESTEST Case #{case_num}")
    runner.registerInfo("Renaming Building > #{model.getBuilding.name}")


    # set timesteps per hour
    timestep = model.getTimestep
    timestep.setNumberOfTimestepsPerHour(4)

    # set shadow calcs
    shadow_calc_freq = model.getShadowCalculation
    shadow_calc_freq.setShadingCalculationUpdateFrequency(1)

    # set ground temps
    ground_temps = model.getSiteGroundTemperatureBuildingSurface
    (1..12).each do |i|
      ground_temps.setTemperatureByMonth(i,10.0)
    end

    # set ground reflectance (default is 0.2 which is what has been used in the past)
    # ground_reflectance = model.getSiteGroundReflectance

    # note: set interior solar distribution fractions isn't needed if E+ auto calcualtes it

    # report final condition of model
    runner.registerFinalCondition("The final model named #{model.getBuilding.name} has #{model.numObjects} objects.")

    return true

  end
  
end

# register the measure to be used by the application
BestestBuildingThermalEnvelopeAndFabricLoad.new.registerWithApplication
