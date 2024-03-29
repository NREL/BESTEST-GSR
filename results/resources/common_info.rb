module BestestResults

  # define cases for section 5.2.3
  def self.populate_common_info(program = "EP")

    hash = {}

    if program == "OS"
      hash[:program_name_and_version] = "OpenStudio 3.4.0"
      hash[:program_version_release_date] = "05/05/2022"
      hash[:program_name_short] = "OS"
    else
      hash[:program_name_and_version] = "EnergyPlus 22.1"
      hash[:program_version_release_date] = "03/29/22"
      hash[:program_name_short] = "E+"
    end

    hash[:results_submission_date] = "05/27/2022"
    hash[:organization] = "National Renewable Energy Laboratory"
    hash[:organization_short] = "NREL"

    return hash

  end

end
