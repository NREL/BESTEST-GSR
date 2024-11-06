module BestestResults

  # common data for spreadsheet headers
  def self.populate_common_info(program = "EP")

    hash = {}

    if program == "OS"
      hash[:program_name_and_version] = "OpenStudio 3.5.1"
      hash[:program_version_release_date] = "12/29/2022"
      hash[:program_name_short] = "OS"
    else
      hash[:program_name_and_version] = "EnergyPlus 22.2"
      hash[:program_version_release_date] = "09/27/2022"
      hash[:program_name_short] = "E+"
    end

    hash[:results_submission_date] = "11/05/2024"
    hash[:organization] = "National Renewable Energy Laboratory"
    hash[:organization_short] = "NREL"

    return hash

  end

end
