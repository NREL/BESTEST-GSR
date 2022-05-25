module BestestCaseVarLib

  # define cases for section 5.2.3
  def self.bestest_5_2_3_case_defs()

    variable_hash_lookup = {}
    variable_hash_lookup['195 - Solid Conduction Test'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.1,
        :ext_ir_emit => 0.1,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => nil,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => true,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['200 - Surface Convection/Infrared Radiation'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.1,
        :ext_ir_emit => 0.1,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => nil,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => true
    ]
    variable_hash_lookup['210 - Interior Infrared Radiation'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.1,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['215 - Exterior Infrared Radiation'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.1,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['220 - In-Depth Series Base Case'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['230 - Infiltration'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 1.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['240 - Internal Gains'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['250 - Exterior Shortwave Absorptance'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.9,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['270 - South Solar Gains'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.9,
        :ext_sw_absorpt => 0.1,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['280 - Cavity Albedo'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.1,
        :ext_sw_absorpt => 0.1,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['290 - South Shading'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.9,
        :ext_sw_absorpt => 0.1,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => 1.0,
        :shade_type => 'H',
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['300 - East/West Window Orientation'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.9,
        :ext_sw_absorpt => 0.1,
        :glass_area => 6.0,
        :orient => 'E,W',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['310 - East/West Shading'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.9,
        :ext_sw_absorpt => 0.1,
        :glass_area => 6.0,
        :orient => 'E,W',
        :shade => 1.0,
        :shade_type => 'H,V',
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['320 - Thermostat'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.9,
        :ext_sw_absorpt => 0.1,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['395 - Solid Conduction Test'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => nil,
        :ext_sw_absorpt => 0.1,
        :glass_area => nil,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => true,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['400 - Opaque Windows with Deadband'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.0,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => nil,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['410 - Infiltration'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 0.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => nil,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['420 - Internal Gains'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => nil,
        :ext_sw_absorpt => 0.1,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['430 - Exterior Shortwave Absorptance'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => nil,
        :ext_sw_absorpt => 0.6,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['440 - Cavity Albedo'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.1,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['450 - Constant Interior and Exterior Surface Coefficients'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :constant_ext_surf_coef => true,
        :constant_int_surf_coef => true
    ]
    variable_hash_lookup['460 - Constant Interior Surface Coefficients'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :constant_ext_surf_coef => false,
        :constant_int_surf_coef => true
    ]
    variable_hash_lookup['470 - Constant Exterior Surface Coefficients'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :constant_ext_surf_coef => true,
        :constant_int_surf_coef => false
    ]
    variable_hash_lookup['600 - Base Case'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['610 - South Shading'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => 1.0,
        :shade_type => 'H',
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['620 - East/West Window Orientation'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 6.0,
        :orient => 'E,W',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['630 - East/West Shading'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 6.0,
        :orient => 'E,W',
        :shade => 1.0,
        :shade_type => 'H,V',
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['640 - Thermostat Setback'] = [
        :htg_set => 'SETBACK',
        :clg_set => 'SETBACK',
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['650 - Night Ventilation'] = [
        :htg_set => 'NONE',
        :clg_set => 27.0,
        :vent => true,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['660 - Glazing Low E'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :glazing_special => 'low-e'
    ]
    variable_hash_lookup['670 - Glazing Single Pane'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :glazing_special => 'single-pane'
    ]
    variable_hash_lookup['680 - Increased Opaque Surface Insulation'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :wall_roof_increased_insulation => true
    ]
    variable_hash_lookup['685 - Thermostat Deadband'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['695 - Thermostat Deadband and Increased Opaque Surface Insulation'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :wall_roof_increased_insulation => true
    ]
    variable_hash_lookup['800 - High-Mass without Solar Gains'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => nil,
        :ext_sw_absorpt => 0.6,
        :glass_area => 0.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => true,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['810 - High-Mass Cavity Albedo'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.1,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['900 - High-Mass Base Case'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['910 - High-Mass South Shading'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => 1.0,
        :shade_type => 'H',
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['920 - High-Mass East/West Window Orientation'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 6.0,
        :orient => 'E,W',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['930 - High-Mass East/West Shading'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 6.0,
        :orient => 'E,W',
        :shade => 1.0,
        :shade_type => 'H,V',
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['940 - High-Mass Thermostat Setback'] = [
        :htg_set => 'SETBACK',
        :clg_set => 'SETBACK',
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['950 - High-Mass Night Ventilation'] = [
        :htg_set => 'NONE',
        :clg_set => 27.0,
        :vent => true,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
      :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    # this case doesn't have any of the typical variables as other cases. Treat as custom solution
    variable_hash_lookup['960 - Sunspace'] = [
        :custom => true,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6
    ]
    variable_hash_lookup['980 - Increased Opaque Surface Insulation'] = [
        :htg_set => 20.0,
        :clg_set => 27.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :wall_roof_increased_insulation => true
    ]
    variable_hash_lookup['985 - Thermostat Deadband'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['995 - Thermostat Deadband and Increased Opaque Surface Insulation'] = [
        :htg_set => 20.0,
        :clg_set => 20.0,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => false,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :wall_roof_increased_insulation => true
    ]
    variable_hash_lookup['600FF'] = [
        :htg_set => "NONE",
        :clg_set => "NONE",
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => true,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['900FF'] = [
        :htg_set => "NONE",
        :clg_set => "NONE",
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => true,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['650FF'] = [
        :htg_set => 'NONE',
        :clg_set => 'NONE',
        :vent => true,
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => true,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]    
    variable_hash_lookup['950FF'] = [
        :htg_set => 'NONE',
        :clg_set => 'NONE',
        :vent => true,
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => true,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false
    ]
    variable_hash_lookup['680FF - Increased Opaque Surface Insulation'] = [
        :htg_set => "NONE",
        :clg_set => "NONE",
        :mass => 'L',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => true,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :wall_roof_increased_insulation => true
    ]    
    variable_hash_lookup['980FF - Increased Opaque Surface Insulation'] = [
        :htg_set => "NONE",
        :clg_set => "NONE",
        :mass => 'H',
        :int_gen => 200.0,
        :infil => 0.5,
        :int_ir_emit => 0.9,
        :ext_ir_emit => 0.9,
        :int_sw_absorpt => 0.6,
        :ext_sw_absorpt => 0.6,
        :glass_area => 12.0,
        :orient => 'S',
        :shade => false,
        :shade_type => nil,
        :ff => true,
        :b1_1_note_01 => false,
        :b1_1_note_02 => false,
        :b1_1_note_03 => false,
        :wall_roof_increased_insulation => true
    ]
    return variable_hash_lookup

  end

  # define cases for section 5.3
  def self.bestest_5_3_case_defs()

    variable_hash_lookup = {}
    variable_hash_lookup['CE100 - Base-Case Building and Mechanical System'] = [
        :series => 'dry zone',
        :int_gen_sensible => 5400.0,
        :int_gen_latent => 0.0,
        :clg_set => 22.2,
        :weather_odb => 46.1,
        :epw => 'CE100A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE110 - Reduced Outdoor Dry-Bulb Temperature'] = [
        :series => 'dry zone',
        :int_gen_sensible => 5400.0,
        :int_gen_latent => 0.0,
        :clg_set => 22.2,
        :weather_odb => 29.4,
        :infil => nil,
        :epw => 'CE110A',
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE120 - Increased Thermostat Setpoint'] = [
        :series => 'dry zone',
        :int_gen_sensible => 5400.0,
        :int_gen_latent => 0.0,
        :clg_set => 26.7,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE130 - Low Part-Load Ratio'] = [
        :series => 'dry zone',
        :int_gen_sensible => 270.0,
        :int_gen_latent => 0.0,
        :clg_set => 22.2,
        :weather_odb => 46.1,
        :epw => 'CE100A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE140 - Reduced Outdoor Dry-Bulb Temperature at Low Part-Load Ratio'] = [
        :series => 'dry zone',
        :int_gen_sensible => 270.0,
        :int_gen_latent => 0.0,
        :clg_set => 22.2,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE150 - Latent Load at High Sensible Heat Ratio'] = [
        :series => 'humid zone',
        :int_gen_sensible => 5400.0,
        :int_gen_latent => 1100.0,
        :clg_set => 22.2,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE160 - Increased Thermostat Setpoint at High Sensible Heat Ratio'] = [
        :series => 'humid zone',
        :int_gen_sensible => 5400.0,
        :int_gen_latent => 1100.0,
        :clg_set => 26.7,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE165 - Variatino fo Thermostat Setpoint and Outdoor Dry-Bulb Temperature at High Sensible Heat Ratio'] = [
        :series => 'humid zone',
        :int_gen_sensible => 5400.0,
        :int_gen_latent => 1100.0,
        :clg_set => 23.3,
        :weather_odb => 40.6,
        :epw => 'CE165A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE170 - Reduced Sensible Load'] = [
        :series => 'humid zone',
        :int_gen_sensible => 2100.0,
        :int_gen_latent => 1100.0,
        :clg_set => 22.2,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE180 - Increased Latent Load'] = [
        :series => 'humid zone',
        :int_gen_sensible => 2100.0,
        :int_gen_latent => 4400.0,
        :clg_set => 22.2,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE185 - Increased Outdoor Dry-Bulb Temperature at Low Sensible Heat Ratio'] = [
        :series => 'humid zone',
        :int_gen_sensible => 2100.0,
        :int_gen_latent => 4400.0,
        :clg_set => 22.2,
        :weather_odb => 46.1,
        :epw => 'CE100A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE190 - Low Part-Load Ratio at Low Sensible Heat Ratio'] = [
        :series => 'humid zone',
        :int_gen_sensible => 270.0,
        :int_gen_latent => 550.0,
        :clg_set => 22.2,
        :weather_odb => 29.4,
        :epw => 'CE110A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE195 - Increased Outdoor Dry-Bulb Temperature at Low Sensible Heat Rato and Low Part-Load Rato'] = [
        :series => 'humid zone',
        :int_gen_sensible => 270.0,
        :int_gen_latent => 550.0,
        :clg_set => 22.2,
        :weather_odb => 46.1,
        :epw => 'CE100A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE200 - Full-Load Test at AHRI Conditions'] = [
        :series => 'full load test at ARI conditions',
        :int_gen_sensible => 6120.0,
        :int_gen_latent => 1817.0,
        :clg_set => 26.7,
        :weather_odb => 35.0,
        :epw => 'CE200A',
        :infil => nil,
        :oa => nil,
        :b1_7_note_a => nil,
        :b1_7_note_b => nil
    ]
    variable_hash_lookup['CE300 - Base Case 15% OA'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE310 - High Latent Load'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'high',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE320 - Infiltration'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 11.558,
        :oa => 0.0,
        :b1_7_note_a => true,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE330 - Outside Air'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 11.558,
        :b1_7_note_a => true,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE340 - Infil/OA Interaction'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 5.779,
        :oa => 5.779,
        :b1_7_note_a => true,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE350 - Thermostat Set Up'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => [25.0,35.0],
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE360 - Undersize'] = [
        :series => 'preliminary',
        :int_gen_sensible => 'high',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :infil => 0.0,
        :epw => 'CE300',
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    # only unique value for CE400 series is case number and descripiton. Will have to use that in measure logic
    variable_hash_lookup['CE400 - Temperature Control'] = [
        :series => 'econmizer',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE410 - Compressor Lockout'] = [
        :series => 'econmizer',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE420 - ODB Limit'] = [
        :series => 'econmizer',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE430 - Enthalpy Control'] = [
        :series => 'econmizer',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE440 - Outdoor Enthalpy Limit'] = [
        :series => 'econmizer',
        :int_gen_sensible => 'mid',
        :int_gen_latent => 'mid',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 1.734,
        :b1_7_note_a => false,
        :b1_7_note_b => false
    ]
    variable_hash_lookup['CE500 - Base Case (0% OA)'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 'mid2',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE510 - High PLR'] = [
        :series => '0 oa',
        :int_gen_sensible => 'high2',
        :int_gen_latent => 'high2',
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE520 - Low EDB 15C'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 'mid2',
        :clg_set => 15.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE522 - Low EDB 20C'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 'mid2',
        :clg_set => 20.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE525 - High EDB'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 'mid2',
        :clg_set => 35.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE530 - Dry Coil'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 0.0,
        :clg_set => 25.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE540 - Dry Coil, Low EDB'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 0.0,
        :clg_set => 15.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]
    variable_hash_lookup['CE545 - Dry Coil, High EDB'] = [
        :series => '0 oa',
        :int_gen_sensible => 'mid2',
        :int_gen_latent => 0.0,
        :clg_set => 35.0,
        :weather_odb => nil,
        :epw => 'CE300',
        :infil => 0.0,
        :oa => 0.0,
        :b1_7_note_a => false,
        :b1_7_note_b => true
    ]

    return variable_hash_lookup

  end

  # define cases for section 5.4
  def self.bestest_5_4_case_defs()

    variable_hash_lookup = {}
    variable_hash_lookup['HE100 - Base-case Building and Mechanical Systems'] = [
        :outdoor_dbt => -30.0,
        :htg_set => 20.0,
        :epw => 'HE100W',
        :capacity => 10.0,
        :ss_eff => 100.0,
        :plr => 1,
        :circ_fan_power => 0.0,
        :circ_fan_type => nil,
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE110 - Efficiency Test'] = [
        :outdoor_dbt => -30.0,
        :htg_set => 20.0,
        :epw => 'HE100W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => 1,
        :circ_fan_power => 0.0,
        :circ_fan_type => nil,
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE120 - Stead Par-Load Test'] = [
        :outdoor_dbt => 0.0,
        :htg_set => 20.0,
        :epw => 'HE120W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => 0.4,
        :circ_fan_power => 0.0,
        :circ_fan_type => nil,
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE130 - No-Load Test'] = [
        :outdoor_dbt => 20.0,
        :htg_set => 20.0,
        :epw => 'HE130W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => 0.0,
        :circ_fan_power => 0.0,
        :circ_fan_type => nil,
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE140 - Periodically Varying Part-Load Test'] = [
        :outdoor_dbt => [-20.0,20.0],
        :htg_set => 20.0,
        :epw => 'HE140W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => [0.0,0.8],
        :circ_fan_power => 0.0,
        :circ_fan_type => nil,
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE150 - Circulating Fan Test'] = [
        :outdoor_dbt => [-20.0,20.0],
        :htg_set => 20.0,
        :epw => 'HE140W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => [0.0,0.8],
        :circ_fan_power => 200.0,
        :circ_fan_type => 'cont',
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE160 - Cycling Circulating Fan Test'] = [
        :outdoor_dbt => [-20.0,20.0],
        :htg_set => 20.0,
        :epw => 'HE140W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => [0.0,0.8],
        :circ_fan_power => 200.0,
        :circ_fan_type => 'cyclic',
        :draft_fan_power => 0.0,
        :draft_fan_type => nil
    ]
    variable_hash_lookup['HE170 - Draft Fan Test'] = [
        :outdoor_dbt => [-20.0,20.0],
        :htg_set => 20.0,
        :epw => 'HE140W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => [0.0,0.8],
        :circ_fan_power => 200.0,
        :circ_fan_type => 'cont',
        :draft_fan_power => 50.0,
        :draft_fan_type => 'cyclic'
    ]
    variable_hash_lookup['HE210 - Realistic Weather Data'] = [
        :outdoor_dbt => 'Varying',
        :htg_set => 20.0,
        :epw => 'HE210W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => [0.0,1.0],
        :circ_fan_power => 200.0,
        :circ_fan_type => 'cyclic',
        :draft_fan_power => 50.0,
        :draft_fan_type => 'cyclic'
    ]
    variable_hash_lookup['HE220 - Setback Thermostat'] = [
        :outdoor_dbt => 'Varying',
        :htg_set => [15.0,20.0],
        :epw => 'HE210W',
        :capacity => 10.0,
        :ss_eff => 80.0,
        :plr => [0.0,1.0],
        :circ_fan_power => 200.0,
        :circ_fan_type => 'cyclic',
        :draft_fan_power => 50.0,
        :draft_fan_type => 'cyclic'
    ]
    variable_hash_lookup['HE230 - Undersized Furnace'] = [
        :outdoor_dbt => 'Varying',
        :htg_set => [15.0,20.0],
        :epw => 'HE210W',
        :capacity => 5.0,
        :ss_eff => 80.0,
        :plr => [0.0,1.0],
        :circ_fan_power => 200.0,
        :circ_fan_type => 'cyclic',
        :draft_fan_power => 50.0,
        :draft_fan_type => 'cyclic'
    ]

    return variable_hash_lookup

  end

  # lookup variables for section 5.2.3 based on case number
  def self.bestest_5_2_3_case_lookup(case_num,runner)

    # lookup case variables
    variable_hash_lookup = BestestCaseVarLib.bestest_5_2_3_case_defs

    # error if can't find expected case
    if variable_hash_lookup[case_num].nil?
      # didn't find expected case
      return false
    else
      variable_hash = variable_hash_lookup[case_num]
    end

    # report out variables in info statements
    runner.registerInfo("Gathering variables for BESTEST Case #{case_num}:")
    variable_hash.first.each do |k,v|
      runner.registerInfo("#{k} = #{v}")
    end

    return variable_hash

  end

  # lookup variables for section 5.3 based on case number
  def self.bestest_5_3_case_lookup(case_num,runner)

    # lookup case variables
    variable_hash_lookup = BestestCaseVarLib.bestest_5_3_case_defs

    # error if can't find expected case
    if variable_hash_lookup[case_num].nil?
      # didn't find expected case
      return false
    else
      variable_hash = variable_hash_lookup[case_num]
    end

    # report out variables in info statements
    runner.registerInfo("Gathering variables for BESTEST Case #{case_num}:")
    variable_hash.first.each do |k,v|
      runner.registerInfo("#{k} = #{v}")
    end

    return variable_hash

  end

  # lookup variables for section 5.4 based on case number
  def self.bestest_5_4_case_lookup(case_num,runner)

    # lookup case variables
    variable_hash_lookup = BestestCaseVarLib.bestest_5_4_case_defs

    # error if can't find expected case
    if variable_hash_lookup[case_num].nil?
      # didn't find expected case
      return false
    else
      variable_hash = variable_hash_lookup[case_num]
    end

    # report out variables in info statements
    runner.registerInfo("Gathering variables for BESTEST Case #{case_num}:")
    variable_hash.first.each do |k,v|
      runner.registerInfo("#{k} = #{v}")
    end

    return variable_hash

  end

end
