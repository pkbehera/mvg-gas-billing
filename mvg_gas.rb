#!/usr/bin/ruby
require 'flat'

#NUM_ARGS = 5
NUM_ARGS = 4
if ARGV.length != NUM_ARGS then
   puts "Requires exactly #{NUM_ARGS} arguemnts"
   #puts 'Usage: ' + $0 + ' KYC_excel_file' + ' Occupancy_excel_file' + ' latest_reading_excel_file' + ' gas_ledger_excel_file' + ' outstanding_excel_file'
   puts 'Usage: ' + $0 + ' KYC_excel_file' + ' Occupancy_excel_file' + ' latest_reading_excel_file' + ' gas_ledger_excel_file'
   abort
end
#if ARGV[0].split(".").last != 'xls' or ARGV[1].split(".").last != 'xls' or ARGV[2].split(".").last != 'xls' or ARGV[3].split(".").last != 'xls' or ARGV[4].split(".").last != 'xls' then
if ARGV[0].split(".").last != 'xls' or ARGV[1].split(".").last != 'xls' or ARGV[2].split(".").last != 'xls' or ARGV[3].split(".").last != 'xls' then
    puts 'Can not process file formats other than .xls (Excel 97-2003), save the files as .xls and retry.'
    abort
end

MVG_ALL_BLOCKS.each do |block|
    block_no = MVG_ALL_BLOCKS.index(block)
    MVG_FLRS_IN_EACH_BLK[block_no].each do |floor|
        MVG_FLTS_IN_EACH_FLR[block_no].each do |flat|
            flat_no = floor.to_s + '%02d' % flat
            block_flat_no = block + flat_no
            if !(MVG_FIRE_REFUSES.include?(block_flat_no) || MVG_IGNORED_FLATS.include?(block_flat_no)) then
                Flat.add_flat(block, flat_no.to_i)
            end
        end
    end
end
count = Flat.get_flat_count
if MVG_TOTAL_NUMBER_OF_FLATS != count then
    puts 'ERROR - the total number of flats ' + MVG_TOTAL_NUMBER_OF_FLATS.to_s + ' does not match with ' + count.to_s + ' records created!'
    abort
end

#Read KYC excel file
puts 'Reading file ' + ARGV[0]
book = Spreadsheet.open ARGV[0]
sheet = book.worksheet 0
row = sheet.row(0)
if row[KYC_BLOCK_COL_NO] != KYC_BLOCK_COL_HEADING or KYC_KYC_COL_HEADING != row[KYC_KYC_COL_NO] then
    puts 'The format of the KYC Excel file ' + ARGV[0] + ' does not seems to be correct, check the file and retry!'
    abort
end
num_flats = 0
sheet.each do |row|
    break if row.join('').empty?
    next if KYC_BLOCK_COL_HEADING == row[KYC_BLOCK_COL_NO]
    block = row[KYC_BLOCK_COL_NO]
    flat = row[KYC_FLAT_COL_NO].to_i
    kyc = false
    subscribed = true
    kyc_s = row[KYC_KYC_COL_NO]
    if kyc_s.nil? or '' == kyc_s.strip then
        subscribed = false
    elsif not kyc_s.nil? and KYC_KYC_OK_STRING_DOWNCASE == kyc_s.strip.downcase then
        kyc = true
    end
    Flat.set_kyc_subscried(block, flat, subscribed, kyc)
    num_flats += 1
end
count = Flat.get_flat_count
if MVG_TOTAL_NUMBER_OF_FLATS != count || MVG_TOTAL_NUMBER_OF_FLATS != num_flats then
    puts 'ERROR - the total number of flats ' + MVG_TOTAL_NUMBER_OF_FLATS.to_s + ' does NOT match with ' + count.to_s + ' records in the KYC file!!'
    abort
end

#Read Occupancy excel file
puts 'Reading file ' + ARGV[1]
book = Spreadsheet.open ARGV[1]
sheet = book.worksheet 0
row = sheet.row(0)
if row[OCC_BLOCK_COL_NO] != OCC_BLOCK_COL_HEADING or OCC_OCC_COL_HEADING != row[OCC_OCC_COL_NO] then
    puts 'The format of the Occupancy Excel file ' + ARGV[1] + ' does not seems to be correct, check the file and retry!'
    abort
end
num_flats = 0
sheet.each do |row|
    break if row.join('').empty?
    next if OCC_BLOCK_COL_HEADING == row[OCC_BLOCK_COL_NO]
    block = row[OCC_BLOCK_COL_NO]
    flat = row[OCC_FLAT_COL_NO].to_i
    occ = true
    occ_s = row[OCC_OCC_COL_NO]
    if not occ_s.nil? and OCC_UNOCC_STRING_DOWNCASE == occ_s.strip.downcase then
        occ = false
    end
    num_flats += 1
    Flat.set_occ(block, flat, occ)
end
count = Flat.get_flat_count
if MVG_TOTAL_NUMBER_OF_FLATS != count || MVG_TOTAL_NUMBER_OF_FLATS != num_flats then
    puts 'ERROR - the total number of flats ' + MVG_TOTAL_NUMBER_OF_FLATS.to_s + ' does NOT match with ' + count.to_s + 'records in the Occupancy file!'
    abort
end

#Read latest gas reading excel file
puts 'Reading file ' + ARGV[2]
count = 0
book = Spreadsheet.open ARGV[2]
sheet = book.worksheet 0
row = sheet.row(0)
if row[READING_BLOCK_COL_NO] != READING_BLOCK_COL_HEADING or READING_CONS_COL_HEADING != row[READING_CONS_COL_NO] then
    puts 'The format of the reading Excel file ' + ARGV[2] + ' does not seems to be correct, check the file and retry!'
    abort
end
sheet.each do |row|
    break if row.join('').empty?
    next if READING_BLOCK_COL_HEADING == row[READING_BLOCK_COL_NO]
    block = row[READING_BLOCK_COL_NO]
    flat = row[READING_FLAT_COL_NO].to_i
    consumed = row[READING_CONS_COL_NO].to_f
    if consumed < 0 then
        puts 'ERROR - NEGATIVE CONSUMPTION FOUND, THERE IS SOMETHING WRONG WITH THE READINGS. CORRECT THE READINGS AND RETRY!'
        abort
    end
    if consumed > 25 then
        puts "WARNING!! HIGH consumtion, #{consumed} m^3 found for #{block}-#{flat}! Continuing bill generation, but review the readings of this flat!"
    end
    Flat.set_consumption(block, flat, consumed)
    count = count + 1
end
puts "Number of readings entered: #{count}"

#Tests
puts 'Checking data sanity...'
Flat.check_sanity

#Read gas ledger excel file
puts 'Reading file ' + ARGV[3]
book = Spreadsheet.open ARGV[3]
sheet = book.worksheet 0
row = sheet.row(0)
if row[LEDGER_BLOCK_COL_NO] != LEDGER_BLOCK_COL_HEADING or row[LEDGER_FLAT_COL_NO] != LEDGER_FLAT_COL_HEADING or row[LEDGER_UNAC_DEBIT_COL_NO] != LEDGER_UNAC_DEBIT_COL_HEADING or row[LEDGER_NO_READING_COUNT_COL_NO] != LEDGER_NO_READING_COUNT_COL_HEADING then
    puts 'The format of the ledger Excel file does not seems to be correct, check the file and retry!'
    abort
end
num_flats = 0
sheet.each do |row|
    break if row.join('').empty?
    next if LEDGER_BLOCK_COL_HEADING == row[LEDGER_BLOCK_COL_NO]
    block = row[LEDGER_BLOCK_COL_NO]
    #Convert flat number to an integer
    flat = row[LEDGER_FLAT_COL_NO].to_i
    sub_used = row[LEDGER_SUB_USED_COL_NO].to_f
    unac_bal = row[LEDGER_UNAC_DEBIT_COL_NO].to_f
    no_reading_cnt = row[LEDGER_NO_READING_COUNT_COL_NO].to_i
    Flat.set_subsidy_unac_debit_no_reading_cnt(block, flat, sub_used, unac_bal, no_reading_cnt)
    num_flats += 1
end
count = Flat.get_flat_count
if MVG_TOTAL_NUMBER_OF_FLATS != count || MVG_TOTAL_NUMBER_OF_FLATS != num_flats then
    puts 'ERROR - the total number of flats ' + MVG_TOTAL_NUMBER_OF_FLATS.to_s + ' does NOT match with ' + count.to_s + 'records in the Ledger file!'
    abort
end

=begin
#Read outstandings excel file
puts 'Reading file ' + ARGV[4]
book = Spreadsheet.open ARGV[4]
sheet = book.worksheet 0
row = sheet.row(1)
if row[OUTSTAND_BLOCK_FLAT_COL_NO] != OUTSTAND_BLOCK_FLAT_COL_HEADING or row[OUTSTAND_BAL_COL_NO] != OUTSTAND_BAL_COL_HEADING then
    puts 'The format of the outstandings Excel file does not seems to be correct, check the file and retry!'
    abort
end
index = 0
sheet.each do |row|
    break if row.join('').empty? and index > 0
    if index < 3 then
        index += 1
        next
    end
    next if OUTSTAND_BLOCK_FLAT_COL_HEADING == row[OUTSTAND_BLOCK_FLAT_COL_NO]
    flat_block = row[OUTSTAND_BLOCK_FLAT_COL_NO].strip
    tokens = flat_block.split("-")
    block = tokens.first.strip
    #Convert flat number to an integer
    flat = tokens.last.to_i
    #A1105 is in the ignore list
    if block == "A" and flat == 1105 then
        index += 1
        next
    end
    outstanding = row[OUTSTAND_BAL_COL_NO].to_f
    Flat.set_outstanding(block, flat, outstanding) if outstanding > 0.0
    index += 1
end
count = Flat.get_flat_count
if MVG_TOTAL_NUMBER_OF_FLATS != count || MVG_TOTAL_NUMBER_OF_FLATS != num_flats then
    puts 'ERROR - the total number of flats ' + MVG_TOTAL_NUMBER_OF_FLATS.to_s + ' does NOT match with ' + count.to_s + 'records in the Ledger file!'
    abort
end
=end

#Now calculate all
puts 'Now calculating debits...'
Flat.process_debits
