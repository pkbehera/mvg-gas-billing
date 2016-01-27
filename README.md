Mont Vert Grande CoOp Cooking Gas Billing
==========================================
Ruby code to generate monthly cooking gas consumption bills for Mont Vert Grande Cooperative Housing Society

**Usage:**
  - Download latest readings from Onesolution and save the file as `Gas_readings_<Mon>_<Year>_downloaded.xls`, e.g. `Gas_readings_Aug_14_downloaded.xls`
  - Extract the KYC & Occupancty columns and save it as `KYC_OCC.xls`
  - Extract the readings provided in the current month, after the last reading of last month, looking for the following errors and correcting them
    - Has anybody provided more than one reading in the current month?
    - Has anybody missed the decimal point in the readings, making the consumption very high?
    - Any other things trying to cheat the software?
  - (optional) Sort all rows in the above files in asscending order, Building first then flat number
  - Use rvm (Ruby Version Manager) to use the system version of Ruby

        rvm use system
  - Go to the directory where the above files are stored
  - Execute the script `mvg_gas.rb` as follows, from the directory where the latest reading file is 

        <path/to/scripts/directory>mvg_gas.rb <optional parameter to redirect console messages to a file> 
