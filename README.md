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
        * integration testing project - bestest_01.xlsx
        * results csv - bestest_os_server_output.csv
        * Std. 140 XLSX - RESULTS5-2A.xlsx
    * Section 5.3 (Space Cooling Equipment Performance Tests)
        * model measure - bestest_space_cooling_equipment_performance_tests
        * reporting measure - bestest_ce_reporting
        * integration testing project - bestest_CE01.xlsx
        * results csv - bestest_os_server_output_ce.csv
        * Std. 140 XLSX - RESULTS5-3A.xlsx and RESULTS5-3B.xlsx
    * Section 5.4 (Space Heating Equipment Performance Test)
        * model measure - bestest_space_heating_equipment_performance_tests
        * reporting measure - bestest_he_reporting
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
* Measures can also be run using the Analysis spreadsheets, thisw allows for quick testing of all test cases.

##### Files used by the Model measures
* measure.rb is the main file that drives the measures. It contains logic to building the model beyond what is included in the tables described below.
* The resoreces folder contains a varity of files used by the measures
    * There is a library of OSM files that contains geometry, schedules, constructinos, and materials. Measures clone these into the seed model.
    * There is a library of EPW files. The measure will assign different EPW files for different test cases.
    * 'besttest_case_var_lib.rb' contains data to create a set of variable values from the case number. Below is a list of tables that are used to create this data.
        * Table B1-1
        * Table B1-2
        * Table B1-5
        * Table B1-7
        * Table B1-8
    * 'besttest_model_methods.rb' contains additional logic and model methods not contained in measure.rb
* The 'tests' folder contains a ruby script to run unit test as well as a seed OSM used by the unit tests.

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
* Add Output Variable requests
 
### Analysis Spreadsheets
 
Use the analysis spreadsheets to run all of the test cases. Currently these are setup as three separate analyses but they could be combined into one. Follow instructions at https://github.com/NREL/OpenStudio-analysis-spreadsheet for using the spreadsheet. You don't need to checkout that repository, all the files you need are already in the OSEP_BESTTEST repository.


##### Integration Testing Files

The files for the Analysis spreadsheet directory are in the 'integration_testing' folder at the top of the repository. Soeme of the files are described below.
* The 'projects' directory contains the three analysis spreadsheet Excel files.
* 'seeds' contains an empty model that serves as the starting point for the workflow
* 'weather' contains a weather file as a starting point but this is replaced by the model measure with a test case specific weather file.

##### Variables Tab

The variables tab for each project has two meaasures. The model measure has a single argument for the case number. There are no arguments for the reporting measure. Sameling is setup to each case is run.

##### Outputs Tab

For each 'runner.RegisterValue' that is used in the reporting measure, a row has to be added to this tab with the 
'Export' column value set to 'TRUE'. As you extend the reporting measure make update this tab or the results won't make it into the downloaded CSV file.

##### OpenStudio Cloud Management Console

Once the analysis has finished, go to the main page for the analysis and under the 'Downloads' section click 'CSV (Results)' to download the CSV file. Then copy it into the 'results' directory and rename it to match the name of the previous section specific CSV file.
 
### Results
 
The 'results' folder at the top of the repository is used to pouplate the 'YourData' tab of the ASHRAE Standard 140-2014 XLSX files. The spreadsheet is setup to display this data on all of the charts.

##### Files in 'results' folder.
* The three CSV files with 'runner.RegisterValue' data are copied here from the server.
* There are three scrips, one for each of the three sections (Enelope, Cooling, Heating).
* The 'resources' folder contains raw Excel files from ASHRAE Standard 140-2014.
* Copies of these Excel files are created at the top of the 'results' folder. There are four files here because the Coooling script generates two Excel files.
* There is an extra copy of each resulting Excel file with "_OS" at the end of the name. These contain the same results but have program infomraiton for OpenStudio instead of EnergyPlus.

##### How the reporting script works.
* Raw Excel File opened and copied.
* CSV resutls loaded and converted to a hash.
* Data from results hash is used to edit the copied Excel file.
* Excel file is saved when changes are done.
* This should be re-run anytime a new Analysis Spreadsheet is run.
* The script is setup to run from the 'results' directory.
* Sometimes Excel Charts won't update unless you copy and paste data in "YourData" worksheet after running the script and opening the file; maybe Excel does not always recognize when file changed via RubyXL?

### Next Steps

* Add to EnergyPlus integration testing
    * Branches can be checked for unexpected changes in BESTEST results
* Extend tests to additional ASHRAE Standard 140-2014 sections beyond 179D requireemnts.