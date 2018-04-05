Building Energy Simulation Test - Generation Simulation and Reporting (BESTEST-GSR)
=============

## The purpose of this repository is to generate BESTEST test case models, run simulations, and populate data for ASHRAE Standard 140 reporting spreadsheets for EnergyPlus based whole building simulation tools.

### Supported tools
The default IDF generation is based on the OpenStudio CLI, but the workflow supports a 'bring your own IDF' use case. The scripts on this repository should work on Mac, Windows, and Linux.

### Dependencies

* Install [OpenStudio 2.5.0](https://www.openstudio.net/downloads)
    * make sure command line can recognize the 'openstudio' command
    * optionally can use other 2.x versions of OpenStudio
* Install [Ruby](https://www.ruby-lang.org/en/) on your system if it isn't already setup.
    * 2.2.4 was used for development but other versions may work
    * Since OpenStudio has it's own embedded Ruby, which is used for running measures, you don't necessarily have to use a version of Ruby supported by OpenStudio.
* Install [RubyXL](https://rubygems.org/gems/rubyXL) ruby gem
    * This is used to modify Microsoft Excel spreadsheets
* Install [Parallel](https://rubygems.org/gems/parallel/versions/1.11.2) ruby gem
    # This allows the CLI to run simulations in parallel

### Steps to Run BESTEST test cases
* Run 'run_all_generate_reports.rb' from the command line script while in the top level of the repository.
    * This should generate IDF files, run simulations, and populate Excel files.
* View resulting files
    * The Summary Excel files are saved to the 'results' directory.
        * 'RESULTS5-#.xlsx' and 'RESULTS5-#_OS.xlsx' have the same content
    * zip file for each datapoint are saved to 'results/bestest_zips'
    
 
### Overview

Currently not all ASHRAE Standard 140-2014 test cases are made. Only test cases required for 179D are created. The following sections are modeled. Under each section is a list of measures, and the Excel file that high level results are created for..
    
* Section 5.2 (Building Thermal Envelope and Fabric Load Tests) except for Section 5.2.4 (Ground Coupling)
    * model measure - bestest_building_thermal_envelope_and_fabric_load_tests
    * reporting measure - bestest_building_thermal_envelope_and_fabric_load_reports
        * Summary of post-processing of SQL data in the measure:
        * Unit conversion (typical)
        * Processing data for specific dates and times (typical)
        * creating annual bins by temperature from hourly data for ff_temp_bins_section method
        * compiling min, max, and average values as needed for ff_temp_bins_section method
    * Std. 140 XLSX - RESULTS5-2A.xlsx
* Section 5.3 (Space Cooling Equipment Performance Tests)
    * model measure - bestest_space_cooling_equipment_performance_tests
    * reporting measure - bestest_ce_reporting
        * Summary of post-processing of SQL data in the measure:
        * Unit Conversion (typical)
        * Processing data for specific dates and times (typical)
        * Calculating COP (mean_cop) using Unitary System Total Cooling Rate and Air System Electric Energy
        * Calculating COP2 (mean_cop_2 and cop_2_24) using Cooling Coil Total Cooling Rate and Air System DX Cooling Coil Electric Energy
        * compiling min, max, and average values as needed
    * Std. 140 XLSX - RESULTS5-3A.xlsx and RESULTS5-3B.xlsx
* Section 5.4 (Space Heating Equipment Performance Test)
    * model measure - bestest_space_heating_equipment_performance_tests
    * reporting measure - bestest_he_reporting
        * Summary of post-processing of SQL data in the measure:
        * Unit Conversion (typical)
        * Calculating average average_fuel_consumption in m^3/sec using formula from section 6.4.1.3
    * Std. 140 XLSX - RESULTS5-4.xlsx


### Measures
 
The measures are contained in the "measures" directory at the top level of the repository. 

##### Running the measures
* Each measure has one or more unit tests
* Measures are run using the OpenStudio CLI and the 'workflow.osw' files described in the 'Workflow' section.
* Measures can also be run as part of the included PAT projects.

##### Files used by the Model measures
* measure.rb is the main file that drives the measures. It contains logic to building the model beyond what is included in the tables described below.
* The resources folder contains a varity of files used by the measures
    * There is a library of OSM files that contains geometry, schedules, constructions, and materials. Measures clone these into the seed model.
    * There is a library of EPW files. The measure will assign different EPW files for different test cases.
    * 'besttest_case_var_lib.rb' contains data to create a set of variable values from the case number. Below is a list of tables that are used to create this data.
        * Table B1-1
        * Table B1-2
        * Table B1-5
        * Table B1-7
        * Table B1-8
    * 'besttest_model_methods.rb' contains additional logic and model methods not contained in measure.rb
* The 'tests' folder contains a ruby script to run unit test as well as a seed OSM used by the unit tests.
* the 'report.html' file generated by the measures can be ignored, the relevent data is saved into the 'out.osw'.
    * The envelope reporting measure has some content in the HTML that was used for measure development, but it not used at all in this workflow.

##### Model articulation steps
* besttest_case_var_lib.rb is loaded
* Case number is mapped ot set of variables
* Adjust simulation settings
* Assign EPW file
* Lookup and load envlope, and make adjustments as needed
* Load bestest_resources.osm.
    * This is a global resource file for all non geometry OSM resources.
* Add construction sets and alter constructions
* Gather schedules that may be required.
* Add Internal Loads
* Add Infiltration
* Add Thermostats
* Add HVAC
* Rename the building
* Add Output Variable requests (These are added by the reporting measure prior to simulation)
 
### Integration Testing Files

#### Workflow

There is a subdirectory for each BESTEST test case. Each directory has a unique 'workflow.osw' file that used to generate the OSM, IDF file, and the simulation results. See link for documentation on the OpenStudio CLI. http://nrel.github.io/OpenStudio-user-documentation/reference/command_line_interface/

There are two sample workflows in the 'workflow' directory, both generate test case 600EN. One articulates the model with OpenStudio and runs an OpenStudio reporting measure. The other imports an externally generated IDF file and run the same OpenStudio reporting measure as the osm generated test case. The only change the reporting measure makes to the IDF is to inject output variable requests needed by the reporting measure. Below is a list of files that are generated when the CLI is used to run an OSW.
* workflow.osw is the only file currently committed ot the repository. The file identifies the seed model, weather file, location of measures and other resources, and the workflow steps that are defined. It can be viewed and edited with a text editor.
* The run directory contains the final OSM, IDF, and simulation results from EnergyPlus.
* The reports directory contains html files generated by reporting measures or EnergyPlus.
* out.osw is generated after running a simulation with the CLI. It is a copy of 'workflow.osw' overloaded with logging from the measure and simulation runs. The data in 'step_values' can be used to populate the results.csv file.
* Only the two .osw files are saved to the repository, however the 'results' directory contains the OSM file, IDF file, and simulation results.
* If you run the BESTEST test cases using 'run_all_generate_reports.rb' it will clean out these directories after using what it needs. You can commment out the script lines that do this, or you can manually run an OSW with the CLI if you want to inspect files not found in the zip file.

Note: the current reporting measures are setup for a specific system confirguration, and in some cases, are expecting objects of specific names. If externally generated IDF files don't match the expected structure and names then the reporting measures, or a copy of it, will need to be altered. The reporting measures could be made more robust so they don't rely on specific names, and support multiple possible system configurations.

#### Parametric Analysis Tool (PAT)

_The PAT projects are no longer used for production but are provided for reference, and for use with manual runs. The current production workflow uses the OpenStudio CLI to directly run 'workflow.osw' files in the 'Workflow' directory._

The Parametric Analysis Tool projects are in the 'integration_testing/pat' folder. Some of the files are described below. See link for PAT documentation. http://nrel.github.io/OpenStudio-user-documentation/reference/parametric_analysis_tool_2/
* There are three separate algorithmic projects for envelope, heating, and cooling. 
* Additionally there is a single manual analysis project with all test cases combined
* The manual PAT project was used to pre-generate the 'osw' files in the 'workflow' directory. If new tests or added or for any reason 'osw' files need to be regenerated, this PAT project can be used.

##### Contents of each PAT project
* pat.json is the PAT project. This is the input file generated by the PAT GUI. It in turn generates the analysis json and analysis zip files.
* measures contains measures used by the analysis. These should be updated from the top level measure directory. The PAT application can manage this.
* seeds contains seed models used by the analysis
* weather contains weather files used by the analysis. Note that some measures may contain their own weather files beyond what is included here.
* project_name.json is the analysis json
* project_name.zip contains all the files needed for the analysis json to run the analysis. 
 
### Results
The 'results' folder at the top of the repository is used to populate the 'YourData' tab of the ASHRAE Standard 140-2014 XLSX files. The spreadsheet is setup to display this data on all of the charts.

##### Files in 'results' folder.
* The is a file nameed 'workflow_results.csv' files with 'runner.RegisterValue' data are copied here from the out.osw files from individual test cases.
* There are three scrips, one for each of the three sections (Enelope, Cooling, Heating).
* The 'resources' folder contains raw Excel files from ASHRAE Standard 140-2014.
* Copies of these Excel files are created at the top of the 'results' folder. There are four files here because the Coooling script generates two Excel files.
* There is an extra copy of each resulting Excel file with "_OS" at the end of the name. These contain the same results but have program information for OpenStudio instead of EnergyPlus.

##### How the reporting script works.
* Raw Excel File opened and copied.
* CSV results loaded and converted to a hash.
* Data from results hash is used to edit the copied Excel file.
* Excel file is saved when changes are done.
* This should be re-run anytime simulations are re-ran. 'run_all_generate_reports.rb' will do this for you.
* The script is setup to run from the 'results' directory.

### Next Steps

* Add to EnergyPlus integration testing
    * Branches can be checked for unexpected changes in BESTEST results
* Extend tests to additional ASHRAE Standard 140-2014 sections beyond 179D requirements.
* Better multiple organization support. Specifically make it easier to maintain running tests across multiple modeling tools from different organizations.
    * Each organization would be responsible for running their own tests, but  as much as the code base as possible can be shared and maintained together.
* Publish high level results to something more dynamic than Excel. Specific interests outlined below.
    * Create reporting with most current version of all tools shown at the same time.
        * This repository is specifically setup for EnergyPlus tools, but this task could fall outside of this repository to support all 179d submissions, with a goal thathigh level data for different tools is easily accessible by the modeling community.
    * Create reporting of the most current version of a tool against some number of prior releases of the same tool.