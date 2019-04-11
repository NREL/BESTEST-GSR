module BestestResults

  # define cases for section 5.2.3
  def self.populate_common_info(program = "EP")

    hash = {}

    if program == "OS"
      hash[:program_name_and_version] = "OpenStudio 2.8.0"
      hash[:program_version_release_date] = "4/12/2019"
      hash[:program_name_short] = "OS"
    else
      hash[:program_name_and_version] = "EnergyPlus 9.1.0"
      hash[:program_version_release_date] = "3/31/2019"
      hash[:program_name_short] = "E+"
    end

    hash[:results_submission_date] = "04/15/2019"
    hash[:organization] = "National Renewable Energy Laboratory"
    hash[:organization_short] = "NREL"

    return hash

  end

end
