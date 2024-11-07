require 'erb'
require 'json'

require File.expand_path("../../shared_resources/os_lib_reporting_bestest", File.dirname(__FILE__))

# start the measure
class BestestBuildingThermalEnvelopeAndFabricLoadReporting < OpenStudio::Measure::ReportingMeasure
  # define the name that a user will see, this method may be deprecated as
  # the display name in PAT comes from the name field in measure.xml
  def name
    return "Bestest Building Thermal Envelope and Fabric Load Reporting"
  end
  # human readable description
  def description
    return "Simple example of modular code to create tables and charts in OpenStudio reporting measures. This is not meant to use as is, it is an example to help with reporting measure development."
  end
  # human readable description of modeling approach
  def modeler_description
    return "This measure uses the same framework and technologies (bootstrap and dimple) that the standard OpenStudio results report uses to create an html report with tables and charts. Download this measure and copy it to your Measures directory using PAT or the OpenStudio application. Then alter the data in os_lib_reporting_custom.rb to suit your needs. Make new sections and tables as needed."
  end
  def possible_sections

    # methods for sections in order that they will appear in report
    result = []

    # instead of hand populating, any methods with 'section' in the name will be added in the order they appear
    all_setions =  OsLib_Reporting_Bestest.methods(false)
    all_setions.each do |section|
      next if not section.to_s.include? 'section'
      result << section.to_s
    end

    result
  end

  # define the arguments that the user will input
  def arguments(model)
    args = OpenStudio::Measure::OSArgumentVector.new

    # this measure does not require any user arguments, return an empty list

    return args
  end

  # add any outout variable requests here
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
    hourly_variables << 'Zone Mean Air Temperature'

    if !bldg_name.include? 'FF' # based on case 600FF
      hourly_variables << 'Zone Air System Sensible Heating Energy'
      hourly_variables << 'Zone Air System Sensible Cooling Energy' # not sure why 630,640,650 dont' have anything below here

      # getsite variables for subset of cases
      if bldg_name.include? "600"
        hourly_variables << 'Site Sky Temperature'
      end

      # get surface variables for subset of cases
      if bldg_name.include? "600"
        hourly_variables << 'Surface Outside Face Sunlit Area'
        hourly_variables << 'Surface Outside Face Sunlit Fraction'
        hourly_variables << 'Surface Outside Face Incident Solar Radiation Rate per Area'
      end

      # get windows variables for subset of cases
      name_test = bldg_name.gsub('BESTEST Case ','')[0..2] # change logic if FF case added that has more characters
      hourly_win_cases = ['600','610','620','630','660','670','900','910','920','930']
      if hourly_win_cases.include?(name_test)
        hourly_variables << 'Surface Window Transmitted Solar Radiation Rate'
        hourly_variables << 'Surface Window Transmitted Beam Solar Radiation Rate'
        hourly_variables << 'Surface Window Transmitted Diffuse Solar Radiation Rate'
        hourly_variables << 'Surface Window Transmitted Solar Radiation Energy'
        hourly_variables << 'Surface Window Transmitted Beam Solar Radiation Energy'
        hourly_variables << 'Surface Window Transmitted Diffuse Solar Radiation Energy'
      end

      # get zone windows variables for subset of cases
      if hourly_win_cases.include?(name_test)
        hourly_variables << 'Zone Windows Total Transmitted Solar Radiation Rate'
      end

    end
    hourly_variables.each do |variable|
      result << OpenStudio::IdfObject.load("Output:Variable,,#{variable},hourly;").get
    end

    # add in monthly variables (needed for OpenStudio 3.7 but not 3.8 and later)
    category_strs = []
    OpenStudio::EndUseCategoryType.getValues.each do |category_type|
      category_str = OpenStudio::EndUseCategoryType.new(category_type).valueDescription
      category_strs << category_str.gsub(" ","")
    end
    monthly_array = ['Output:Table:Monthly']
    monthly_array << "Building Energy Performance - District Heating Water"
    monthly_array << '2'
    category_strs.each do |category_string|
      monthly_array << "#{category_string}:DistrictHeatingWater"
      monthly_array << 'SumOrAverage'
    end
    result << OpenStudio::IdfObject.load("#{monthly_array.join(',')};").get
    monthly_array = ['Output:Table:Monthly']
    monthly_array << "Building Energy Performance - District Heating Water Peak Demand"
    monthly_array << '2'
    monthly_array << "DistrictHeatingWater:Facility"
    monthly_array << "Maximum"
    category_strs.each do |category_string|
      monthly_array << "#{category_string}:DistrictHeatingWater"
      monthly_array << 'ValueWhenMaximumOrMinimum'
    end
    result << OpenStudio::IdfObject.load("#{monthly_array.join(',')};").get

    result
  end

  def outputs
    result = OpenStudio::IdfObjectVector.new

    # items here are out date for Std 140 2020. Since these are not used now, since we are in PAT it isn't useful to mainain
    # results.csv is generated by a post proessing script and it grabs all runner.registerValues

    return result

  end

  # define what happens when the measure is run
  def run(runner, user_arguments)
    super(runner, user_arguments)

    # get sql, model, and web assets
    setup = OsLib_Reporting_Bestest.setup(runner)
    unless setup
      return false
    end
    model = setup[:model] # no data from model used, just needed to call specific methods
    # workspace = setup[:workspace]
    sql_file = setup[:sqlFile]
    web_asset_path = setup[:web_asset_path]

    # reporting final condition
    runner.registerInitialCondition('Gathering data from EnergyPlus SQL file.')

    # pass measure display name to erb
    @name = name

    # create a array of sections to loop through in erb file
    @sections = []

    # generate data for requested sections
    sections_made = 0
    possible_sections.each do |method_name|

      begin
        section = false
        eval("section = OsLib_Reporting_Bestest.#{method_name}(model,sql_file,runner,false)")
        display_name = eval("OsLib_Reporting_Bestest.#{method_name}(nil,nil,nil,true)[:title]")
        if section
          @sections << section
          sections_made += 1
          # look for emtpy tables and warn if skipped because returned empty
          section[:tables].each do |table|
            if not table
              #runner.registerWarning("A table in #{display_name} section returned false and was skipped.")
              #section[:messages] = ["One or more tables in #{display_name} section returned false and was skipped."]
            end
          end
        else
          # runner.registerWarning("#{display_name} section returned false and was skipped.")
          section = {}
          section[:title] = "#{display_name}"
          section[:tables] = []
          section[:messages] = []
          section[:messages] << "#{display_name} section returned false and was skipped."
          # @sections << section
        end
      rescue => e
        display_name = eval("OsLib_Reporting_Bestest.#{method_name}(nil,nil,nil,true)[:title]")
        if display_name == nil then display_name == method_name end
        #runner.registerWarning("#{display_name} section failed and was skipped because: #{e}. Detail on error follows.")
        #runner.registerWarning("#{e.backtrace.join("\n")}")

        # add in section heading with message if section fails
        section = eval("OsLib_Reporting_Bestest.#{method_name}(nil,nil,nil,true)")
        section[:messages] = []
        section[:messages] << "#{display_name} section failed and was skipped because: #{e}. Detail on error follows."
        section[:messages] << ["#{e.backtrace.join("\n")}"]
        #@sections << section

      end

    end

    # read in template
    html_in_path = "#{File.dirname(__FILE__)}/resources/report.html.erb"
    if File.exist?(html_in_path)
      html_in_path = html_in_path
    else
      html_in_path = "#{File.dirname(__FILE__)}/report.html.erb"
    end
    html_in = ''
    File.open(html_in_path, 'r') do |file|
      html_in = file.read
    end

    # configure template with variable values
    renderer = ERB.new(html_in)
    html_out = renderer.result(binding)

    # write html file
    html_out_path = './report.html'
    File.open(html_out_path, 'w') do |file|
      file << html_out
      # make sure data is written to the disk one way or the other
      begin
        file.fsync
      rescue
        file.flush
      end
    end

    # closing the sql file
    sql_file.close

    # reporting final condition
    runner.registerFinalCondition("Generated report with #{sections_made} sections to #{html_out_path}.")

    true
  end # end the run method
end # end the measure

# this allows the measure to be use by the application
BestestBuildingThermalEnvelopeAndFabricLoadReporting.new.registerWithApplication
