OS_EP_Bestest
=============

BESTEST

This document gives overview of task
https://docs.google.com/document/d/1OOjD81k1JyTmbGo-u2SEfPAduMdXrWaHakjF9Uu4do8/edit?usp=sharing


This folder contains design documents for specific measures.
https://drive.google.com/folderview?id=0BziGk7eItbVVYk5JX3VNcFk0bjQ&usp=sharing


## Workflow to generate BESTEST test cases and populating ASHRAE Standard 140 reporting spreadsheets

### Dependancies

* checkout this repository
* Make sure OpenStudio is accessible from ruby on your system.
* install RubyXL gem (not needed for measures, only for writing to Excel in with reporting script)


### Overview

* Currently not all ASHRAE Standard 140-2014 test cases are made. Only test cases required for 179D are created. The following sections are modeled
    * Section 5.2 (Building Thermal Envelope and Fabric Load Tests) except for Section 5.2.4 (Ground Coupling)
        * model measure - bestest_building_thermal_envelope_and_fabric_load_tests
        * reporting measure - bestest_building_thermal_envelope_and_fabric_load_reports
            * Summary of post-processing of SQL data in the measure:
            * Unit conversion (typical)
            * Processing data for specific dates and times (typcial)
            * creating annual bins by temperature from hourly data for ff_temp_bins_section method
            * compiling min, max, and average values as needed for ff_temp_bins_section method
        * integration testing project - bestest_01.xlsx
        * results csv - bestest_os_server_output.csv
        * Std. 140 XLSX - RESULTS5-2A.xlsx
    * Section 5.3 (Space Cooling Equipment Performance Tests)
        * model measure - bestest_space_cooling_equipment_performance_tests
        * reporting measure - bestest_ce_reporting
            * Summary of post-processing of SQL data in the measure:
            * Unit Conversion (typical)
            * Processing data for specific dates and times (typcial)
            * Calculating COP (mean_cop) using Unitary System Total Cooling Rate and Air System Electric Energy
            * Calculating COP2 (mean_cop_2 and cop_2_24) using Cooling Coil Total Cooling Rate and Air System DX Cooling Coil Electric Energy
            * compiling min, max, and average values as needed
        * integration testing project - bestest_CE01.xlsx
        * results csv - bestest_os_server_output_ce.csv
        * Std. 140 XLSX - RESULTS5-3A.xlsx and RESULTS5-3B.xlsx
    * Section 5.4 (Space Heating Equipment Performance Test)
        * model measure - bestest_space_heating_equipment_performance_tests
        * reporting measure - bestest_he_reporting
            * Summary of post-processing of SQL data in the measure:
            * Unit Conversion (typical)
            * Calculating average average_fuel_consumption in m^3/sec using formula from section 6.4.1.3
        * integration testing project - bestest_HE01.xlsx
        * results csv - bestest_os_server_output_he.csv
        * Std. 140 XLSX - RESULTS5-4.xlsx
* Each Section listed above has it's own set of measures and scripts. Below is a list of what is used for each section.
    * First an OpenStudio Model measure is used to create the test case and assign a weather file
    * Then an OpenStudio Reporting measure is used to generate 'runner.RegisterValue' objects that will be used downstream.
    * Next there is an OpenStudio Analysis Spreadsheet project that runs the measures and gathers the results from the individual datapoints. 
    * A a CSV file is downloaded with the the 'runner.RegisterValue' log from each datapoint.
    * A local script then runs and using data from the CSV, edits a copy of the raw ASHRAE Standard 140-2014 Excel Spreadsheet.
 
### Measures
 
The measures are contained in the "measures" directory at the top level of the repository.

##### Running the measures
* Each measure has one or more unit tests
* Measrues can also be run directly in the OpenStudio applciation, howeer there is a bug in OpenStudio that prevents weather files from being assigned, so you have to add the weather file manually before running a simulation.

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

##### Model articulation
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

#### Parametric Analysis Tool (PAT)

The Parametric Analysis Tool projects are in the 'integration_testing/pat' folder. Some of the files are described below. See link for PAT documentation. http://nrel.github.io/OpenStudio-user-documentation/reference/parametric_analysis_tool_2/
* There are three separate algorithmic projects for envelope, heating, and cooling. 
* Additionally there is a single manual analysis project with all test cases combined

##### Contents of each PAT project
* pat.json is the PAT project. This is the input file generated by the PAT GUI. It in turn generates the analysis json and analysis zip files.
* measures contains measures used by the analysis. These should be updated from the top level measure directory. The PAT application can manage this.
* seeds contains seed models used by the analysis
* weather contains weather files used by the analysis. Note that some measures may contain their own weather files beyond what is included here.
* project_name.json is the analysis json
* project_name.zip contains all the files needed for the analysis json to run the analysis.

##### OpenStudio Cloud Management Console

Once the analysis has finished, go to the main page for the analysis and under the 'Downloads' section click 'CSV (Results)' to download the CSV file. Then copy it into the 'results' directory and rename it to match the name of the previous section specific CSV file.

#### Workflow

This directory contains example workflows that can be run with the OpenStudio CLI to generate test cases and run simulations. See link for documentation on the OpenStudio CLI. http://nrel.github.io/OpenStudio-user-documentation/reference/command_line_interface/

There are two sample workflows in the 'workflow' directory, both generate test case 600EN. One articulations the model with OpenStudio and runs an OpenStudio reporting measure. The other imports an externally generated IDF file and run the same OpenStudio reporting measure as the osm generated test case. The only change the reporting measure makes to the IDF is to inject output variable requests needed by the reporting measure. Below is a list of files that are generated when the CLI is used to run an OSW.
* workflow.osw is the only file currently committed ot the repository. The can identifies the seed model, weather file, location of measures and other resoruces, and the workflow steps that are defined. It can be viewed and edited with a text editor.
* The run directory contains the final OSM, IDF, and simulation results from EnergyPlus.
* The reports directory contains html files generated by reporting measures or EnergyPlus.
* out.osw is geneated after running a simulation with the CLI. It is a copy of 'workflow.osw' overloaded with logging from the measure and simulation runs. The data in 'step_values' can be used to populate the results.csv file. 

Note: the current reporting measures are setup for a specific system confirguration, and in some cases, are expecting objects of specific names. If externally generated IDF files don't match the expected structurea and names then the reporting measures, or a copy of it, will need to be altered. The reporting measures could be made more robust so they don't rely on specific names, and support mutliple possible system configurations.
 
### Results
The 'results' folder at the top of the repository is used to populate the 'YourData' tab of the ASHRAE Standard 140-2014 XLSX files. The spreadsheet is setup to display this data on all of the charts.

##### Files in 'results' folder.
* The three CSV files with 'runner.RegisterValue' data are copied here from the server.
* There are three scrips, one for each of the three sections (Enelope, Cooling, Heating).
* The 'resources' folder contains raw Excel files from ASHRAE Standard 140-2014.
* Copies of these Excel files are created at the top of the 'results' folder. There are four files here because the Coooling script generates two Excel files.
* There is an extra copy of each resulting Excel file with "_OS" at the end of the name. These contain the same results but have program infomraiton for OpenStudio instead of EnergyPlus.

##### How the reporting script works.
* Raw Excel File opened and copied.
* CSV results loaded and converted to a hash.
* Data from results hash is used to edit the copied Excel file.
* Excel file is saved when changes are done.
* This should be re-run anytime a new Analysis Spreadsheet is run.
* The script is setup to run from the 'results' directory.
* Sometimes Excel Charts won't update unless you copy and paste data in "YourData" worksheet after running the script and opening the file; maybe Excel does not always recognize when file changed via RubyXL?

### Next Steps

* Add to EnergyPlus integration testing
    * Branches can be checked for unexpected changes in BESTEST results
* Extend tests to additional ASHRAE Standard 140-2014 sections beyond 179D requireemnts.
* Setup post processing script to work from either results.csv or collection of individual out.osw files.
    * Most direct way to do this is to add in code to create results.csv from individual osw files, then call the existing code as is. This will allow manual local runs using PAT or CLI.
* Better support for multiple organization support.
    * Add arguments to scripts to point to location of results and general information. Also good time to combine scripts under a single parent script.