# see the URL below for information on how to write OpenStudio measures
# http://nrel.github.io/OpenStudio-user-documentation/reference/measure_writing_guide/

# start the measure
class ReplaceIdf < OpenStudio::Ruleset::WorkspaceUserScript

  # human readable name
  def name
    return "Replace Idf"
  end

  # human readable description
  def description
    return "Replace this text with an explanation of what the measure does in terms that can be understood by a general building professional audience (building owners, architects, engineers, contractors, etc.).  This description will be used to create reports aimed at convincing the owner and/or design team to implement the measure in the actual building design.  For this reason, the description may include details about how the measure would be implemented, along with explanations of qualitative benefits associated with the measure.  It is good practice to include citations in the measure if the description is taken from a known source or if specific benefits are listed."
  end

  # human readable description of modeling approach
  def modeler_description
    return "Replace this text with an explanation for the energy modeler specifically.  It should explain how the measure is modeled, including any requirements about how the baseline model must be set up, major assumptions, citations of references to applicable modeling resources, etc.  The energy modeler should be able to read this description and understand what changes the measure is making to the model and why these changes are being made.  Because the Modeler Description is written for an expert audience, using common abbreviations for brevity is good practice."
  end

  # define the arguments that the user will input
  def arguments(workspace)
    args = OpenStudio::Ruleset::OSArgumentVector.new

    # argument for external_idf_name
    external_idf_name = OpenStudio::Ruleset::OSArgument.makeStringArgument("external_idf_name", true)
    external_idf_name.setDisplayName("External IDF File Name")
    external_idf_name.setDescription("Name of the model to replalace current model. This is the filename with the extension (e.g. MyModel.idf). Optionally this can inclucde the full file path, but for most use cases should just be file name.")
    args << external_idf_name

    return args
  end 

  # define what happens when the measure is run
  def run(workspace, runner, user_arguments)
    super(workspace, runner, user_arguments)

    # use the built-in error checking 
    if !runner.validateUserArguments(arguments(workspace), user_arguments)
      return false
    end

    # assign the user inputs to variables
    external_idf_name = runner.getStringArgumentValue("external_idf_name", user_arguments)

    # find external model
    idf_file = runner.workflow.findFile(external_idf_name)
    if idf_file.is_initialized
      external_idf_name = idf_file.get.to_s
    else
      runner.registerError("Did not find #{external_idf_name} in paths described in OSW file.")
      runner.registerInfo("Looked for #{external_idf_name} in the following locations")
      runner.workflow.absoluteFilePaths.each do |path|
        runner.registerInfo("#{path}")
      end
      return false
    end

    # get model from path and error if empty
    external_workspace = OpenStudio::Workspace::load(OpenStudio::Path.new(external_idf_name))
    if external_workspace.empty?
      runner.registerError("Cannot load #{external_idf_name}")
      return false
    end
    external_workspace = external_workspace.get

    # report initial condition of model
    runner.registerInitialCondition("The initial IDF file had #{workspace.objects.size} objects.")

    runner.registerInfo("Loading #{external_idf_name}")

    # remove existing objects from model
    handles = OpenStudio::UUIDVector.new
    workspace.objects.each do |obj|
      handles << obj.handle
    end
    workspace.removeObjects(handles)

    # add new file to empty model
    workspace.addObjects( external_workspace.toIdfFile.objects )

    # report final condition of model
    runner.registerFinalCondition("The final IDF file had #{workspace.objects.size} objects.")
 
  end

end 

# register the measure to be used by the application
ReplaceIdf.new.registerWithApplication
