module BestestResults

  # common data for spreadsheet headers
  def self.populate_common_info(program = "EP")

    hash = {}

    if program == "OS"
      hash[:program_name_and_version] = "OpenStudio 3.6.0"
      hash[:program_version_release_date] = "05/09/2023"
      hash[:program_name_short] = "OS"
    else
      hash[:program_name_and_version] = "EnergyPlus 23.1"
      hash[:program_version_release_date] = "03/28/2023"
      hash[:program_name_short] = "E+"
    end

    hash[:results_submission_date] = "05/09/2023"
    hash[:organization] = "National Renewable Energy Laboratory"
    hash[:organization_short] = "NREL"

    return hash

  end

end
