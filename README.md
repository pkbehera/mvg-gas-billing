mvg-gas-billing
===============

Ruby code for generating monthly cooking gas consumption bills for Mont Vert Grande Cooperative Housing Society

**Usage:**
  - Download latest readings from Onesolution and save the file as `Gas_readings_<Mon>_<Year>_downloaded.xls`, e.g. `Gas_readings_Aug_14_downloaded.xls`

  - Extract the KYC column and save it as `KYC.xls`

  - Extract the Occupancty column and save it as `OCC.xls`

  - Extract the readings provided in the current month, after the last reading of last month, looking for the following errors and correcting them

    - Has anybody provided more than one reading in the current month?

    - Has anybody missed the decimal point in the readings, making the consumption very high?

  - (optional) Sort all the above files in asscending order, Building first then flat number

  - Use rvm (Ruby Version Manager) to use the system version of Ruby

        rvm use system

  - Execute the script `mvg_gas.rb` as follows

        mvg_gas.rb <KYC_excel> <OCC_excel> <latest_reading_excel> <gas_ledger_excel>
