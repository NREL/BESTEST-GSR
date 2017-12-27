require 'openstudio'
require 'openstudio/ruleset/ShowRunnerOutput'
require 'minitest/autorun'

require_relative '../measure.rb'

require 'fileutils'

class BestestCEReporting_Test < MiniTest::Unit::TestCase

  def is_openstudio_2?
    begin
      workflow = OpenStudio::WorkflowJSON.new
    rescue
      return false
    end
    return true
  end

  def model_in_path_default
    return nil # no default input model for this test
  end

  def epw_path_default
    return nil # no default epw for this test
  end

  def run_dir(test_name)
    # always generate test output in specially named 'output' directory so result files are not made part of the measure
    "#{File.dirname(__FILE__)}/output/#{test_name}"
  end

  def model_out_path(test_name)
    "#{run_dir(test_name)}/TestOutput.osm"
  end

  def workspace_path(test_name)
    if is_openstudio_2?
      return "#{run_dir(test_name)}/run/in.idf"
    else
      return "#{run_dir(test_name)}/ModelToIdf/in.idf"
    end
  end

  def sql_path(test_name)
    if is_openstudio_2?
      return "#{run_dir(test_name)}/run/eplusout.sql"
    else
      return "#{run_dir(test_name)}/ModelToIdf/EnergyPlusPreProcess-0/EnergyPlus-0/eplusout.sql"
    end
  end

  def report_path(test_name)
    "#{run_dir(test_name)}/report.html"
  end

  # method for running the test simulation using OpenStudio 1.x API
  def setup_test_1(test_name, epw_path)

    co = OpenStudio::Runmanager::ConfigOptions.new(true)
    co.findTools(false, true, false, true)

    if !File.exist?(sql_path(test_name))
      puts "Running EnergyPlus"

      wf = OpenStudio::Runmanager::Workflow.new("modeltoidf->energypluspreprocess->energyplus")
      wf.add(co.getTools())
      job = wf.create(OpenStudio::Path.new(run_dir(test_name)), OpenStudio::Path.new(model_out_path(test_name)), OpenStudio::Path.new(epw_path))

      rm = OpenStudio::Runmanager::RunManager.new
      rm.enqueue(job, true)
      rm.waitForFinished
    end
  end

  # method for running the test simulation using OpenStudio 2.x API
  def setup_test_2(test_name, epw_path)

    if !File.exist?(sql_path(test_name))
      osw_path = File.join(run_dir(test_name), 'in.osw')
      osw_path = File.absolute_path(osw_path)

      workflow = OpenStudio::WorkflowJSON.new
      workflow.setSeedFile(File.absolute_path(model_out_path(test_name)))
      workflow.setWeatherFile(File.absolute_path(epw_path))
      workflow.saveAs(osw_path)

      cli_path = OpenStudio.getOpenStudioCLI
      cmd = "\"#{cli_path}\" run -w \"#{osw_path}\""
      puts cmd
      system(cmd)
    end
  end

  # create test files if they do not exist when the test first runs
  def setup_test(test_name, idf_output_requests, model_in_path = model_in_path_default, epw_path = epw_path_default)

    if !File.exist?(run_dir(test_name))
      FileUtils.mkdir_p(run_dir(test_name))
    end
    assert(File.exist?(run_dir(test_name)))

    if File.exist?(report_path(test_name))
      FileUtils.rm(report_path(test_name))
    end

    assert(File.exist?(model_in_path))

    if File.exist?(model_out_path(test_name))
      FileUtils.rm(model_out_path(test_name))
    end

    # convert output requests to OSM for testing, OS App and PAT will add these to the E+ Idf
    workspace = OpenStudio::Workspace.new("Draft".to_StrictnessLevel, "EnergyPlus".to_IddFileType)
    workspace.addObjects(idf_output_requests)
    rt = OpenStudio::EnergyPlus::ReverseTranslator.new
    request_model = rt.translateWorkspace(workspace)

    translator = OpenStudio::OSVersion::VersionTranslator.new
    model = translator.loadModel(model_in_path)
    assert((not model.empty?))
    model = model.get
    model.addObjects(request_model.objects)
    model.save(model_out_path(test_name), true)

    if is_openstudio_2?
      setup_test_2(test_name, epw_path)
    else
      setup_test_1(test_name, epw_path)
    end
  end

  def apply_measure_to_model(test_name, args_hash, model_in_path, epw_path)

    # create an instance of the measure
    measure = BestestCeReporting.new

    # create an instance of a runner
    runner = OpenStudio::Ruleset::OSRunner.new

    # get arguments
    arguments = measure.arguments
    argument_map = OpenStudio::Ruleset.convertOSArgumentVectorToMap(arguments)

    # populate argument with specified hash value if specified
    arguments.each do |arg|
      temp_arg_var = arg.clone
      if args_hash[arg.name]
        assert(temp_arg_var.setValue(args_hash[arg.name]))
      end
      argument_map[arg.name] = temp_arg_var
    end

    # only need to run this when osm models are added or updated
=begin
    # need to set last workspace early so proper variables can be grabbed
    # load the test model
    translator = OpenStudio::OSVersion::VersionTranslator.new
    model = translator.loadModel(model_in_path)
    assert((not model.empty?))
    model = model.get
    # forward translate OSM file to IDF file
    ft = OpenStudio::EnergyPlus::ForwardTranslator.new
    workspace = ft.translateModel(model)
    workspace_in_path = model_in_path.gsub(".osm",".idf")
    workspace.save(workspace_in_path,true)
=end

    # get the energyplus output requests, this will be done automatically by OS App and PAT
    runner.setLastEnergyPlusWorkspacePath(OpenStudio::Path.new(model_in_path.gsub(".osm",".idf")))
    idf_output_requests = measure.energyPlusOutputRequests(runner, argument_map)
    assert_operator idf_output_requests.size, :>, 0

    # mimic the process of running this measure in OS App or PAT
    setup_test(test_name,idf_output_requests,model_in_path,epw_path)

    assert(File.exist?(model_out_path(test_name)))
    assert(File.exist?(sql_path(test_name)))
    assert(File.exist?(epw_path))

    # set up runner, this will happen automatically when measure is run in PAT or OpenStudio
    runner.setLastOpenStudioModelPath(OpenStudio::Path.new(model_out_path(test_name)))
    runner.setLastEnergyPlusWorkspacePath(OpenStudio::Path.new(workspace_path(test_name)))
    runner.setLastEpwFilePath(epw_path)
    runner.setLastEnergyPlusSqlFilePath(OpenStudio::Path.new(sql_path(test_name)))

    # delete the output if it exists
    if File.exist?(report_path(test_name))
      FileUtils.rm(report_path(test_name))
    end
    assert(!File.exist?(report_path(test_name)))

    # temporarily change directory to the run directory and run the measure
    start_dir = Dir.pwd
    begin
      Dir.chdir(run_dir(test_name))

      # run the measure
      measure.run(runner, argument_map)
      result = runner.result
      show_output(result)
      assert_equal('Success', result.value.valueName)

    ensure
      Dir.chdir(start_dir)
    end

    # no report.html is written
    # assert(File.exist?(report_path(test_name)))
  end

  def test_case_CE100
    args = {}
    model_in_path = "#{File.dirname(__FILE__)}/CE100_test_output.osm"
    epw_path = "#{File.dirname(__FILE__)}/CE100ATM2.epw"

    apply_measure_to_model(__method__.to_s.gsub('test_',''), args, model_in_path, epw_path)
  end

  def test_case_CE200
    args = {}
    model_in_path = "#{File.dirname(__FILE__)}/CE200_test_output.osm"
    epw_path = "#{File.dirname(__FILE__)}/CE200ATM2.epw"

    apply_measure_to_model(__method__.to_s.gsub('test_',''), args, model_in_path, epw_path)
  end

  def test_case_CE300
    args = {}
    model_in_path = "#{File.dirname(__FILE__)}/CE300_test_output.osm"
    epw_path = "#{File.dirname(__FILE__)}/CE300TM2.epw"

    apply_measure_to_model(__method__.to_s.gsub('test_',''), args, model_in_path, epw_path)
  end

  def test_case_CE400
    args = {}
    model_in_path = "#{File.dirname(__FILE__)}/CE400_test_output.osm"
    epw_path = "#{File.dirname(__FILE__)}/CE300TM2.epw"

    apply_measure_to_model(__method__.to_s.gsub('test_',''), args, model_in_path, epw_path)
  end

  def test_case_CE500
    args = {}
    model_in_path = "#{File.dirname(__FILE__)}/CE500_test_output.osm"
    epw_path = "#{File.dirname(__FILE__)}/CE300TM2.epw"

    apply_measure_to_model(__method__.to_s.gsub('test_',''), args, model_in_path, epw_path)
  end

end
