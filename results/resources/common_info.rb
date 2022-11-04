module BestestResults

  # common data for spreadsheet headers
  def self.populate_common_info(program = "EP")

    hash = {}

    if program == "OS"
      hash[:program_name_and_version] = "OpenStudio 3.5.0"
      hash[:program_version_release_date] = "11/03/2022"
      hash[:program_name_short] = "OS"
    else
      hash[:program_name_and_version] = "EnergyPlus 22.2"
      hash[:program_version_release_date] = "09/27/22"
      hash[:program_name_short] = "E+"
    end

    hash[:results_submission_date] = "11/03/2022"
    hash[:organization] = "National Renewable Energy Laboratory"
    hash[:organization_short] = "NREL"

    return hash

  end

end
