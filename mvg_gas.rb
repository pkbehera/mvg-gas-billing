#!/usr/bin/ruby
scriptDir = File.dirname(__FILE__)
require "#{scriptDir}/flat"
require "date"

cwd = Dir.getwd
workDir = File.basename(cwd)
today = Date.today
reqDir = today.strftime("%b%Y")
curMon = today.strftime("%b")
curYr = today.strftime("%Y")

#Same date in last month
prevMonth = today << 1
preMonthDir = prevMonth.strftime("%b%Y")
preMonth = prevMonth.strftime("%b_%Y")
preMon = prevMonth.strftime("%b")
preMonYr = prevMonth.strftime("%Y")

if ! workDir.eql? reqDir then
   puts "This script should be run from a folder named #{reqDir}"
   abort
end

kycOccFile = "KYC_OCC.xls"
if ! File.exist?(kycOccFile) then
   puts "File #{kycOccFile} does not exist"
   abort
end

rdngFilePat = "gasmeter_readings_{[0-3][0-9]}#{curMon}_to_{[0-3][0-9]}#{curMon}.xls"
rdngFilePat1 = "gasmeter_readings_{[0-3][0-9]}#{preMon}_to_{[0-3][0-9]}#{curMon}.xls"

list = Dir.glob(rdngFilePat)
list += (Dir.glob(rdngFilePat1))
if list.empty? then
   puts "Current month readings file of pattern #{rdngFilePat} or #{rdngFilePat1} does not exist"
   abort
else
   if list.size > 1 then
      puts "There are more than one file of pattern #{rdngFilePat}"
      abort
   end
   rdngFile = list[0]
end

ledgFile = "../../#{preMonYr}/#{preMonthDir}/Gas_Ledger_#{preMonth}.xls"
if ! File.exist?(ledgFile) then
   puts "File #{ledgFile} does not exist"
   abort
end

if ARGV.length > 0 then
   notesFile = "Notes_#{curMon}_#{curYr}.txt"
   puts "Redirecting console output to #{notesFile}"
   $stdout = File.new(notesFile, 'w')
   $stdout.sync = true
end

puts "Current working directory #{cwd}"

'''
#NUM_ARGS = 5
NUM_ARGS = 4
if ARGV.length != NUM_ARGS then
   puts "Requires exactly #{NUM_ARGS} arguemnts"
   #puts "Usage: #{$0} KYC_file latest_reading_file gas_ledger_file outstanding_file"
   puts "Usage: #{$0} KYC_file latest_reading_file gas_ledger_file"
   abort
end

kycOccFile = ARGV[0]
rdngFile = ARGV[1]
ledgFile = ARGV[2]
outsdFile = ARGV[3]
'''

#if kycOccFile.split(".").last != 'xls' or rdngFile.split(".").last != 'xls' or ledgFile.split(".").last != 'xls' or outsdFile.split(".").last != 'xls' then
if kycOccFile.split(".").last != 'xls' or rdngFile.split(".").last != 'xls' or ledgFile.split(".").last != 'xls' then
    puts "Can not process file formats other than .xls (Excel 97-2003), save the files as .xls and retry."
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
puts 'Reading file ' + kycOccFile
book = Spreadsheet.open kycOccFile
sheet = book.worksheet 0
row = sheet.row(0)
if row[KYC_BLOCK_COL_NO] != KYC_BLOCK_COL_HEADING or KYC_KYC_COL_HEADING != row[KYC_KYC_COL_NO] or OCC_OCC_COL_HEADING != row[OCC_OCC_COL_NO] then
    puts 'The format of the KYC/Occupancy Excel file ' + kycOccFile + ' does not seems to be correct, check the file and retry!'
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
    #Read Occupancy flag
    occ = true
    occ_s = row[OCC_OCC_COL_NO]
    if not occ_s.nil? and OCC_UNOCC_STRING_DOWNCASE == occ_s.strip.downcase then
        occ = false
    end
    Flat.set_occ(block, flat, occ)
    num_flats += 1
end
count = Flat.get_flat_count
if MVG_TOTAL_NUMBER_OF_FLATS != count || MVG_TOTAL_NUMBER_OF_FLATS != num_flats then
    puts 'ERROR - the total number of flats ' + MVG_TOTAL_NUMBER_OF_FLATS.to_s + ' does NOT match with ' + count.to_s + ' records in the KYC file!!'
    abort
end

#Read latest gas reading excel file
puts 'Reading file ' + rdngFile
count = 0
book = Spreadsheet.open rdngFile
sheet = book.worksheet 0
row = sheet.row(0)
if row[READING_BLOCK_COL_NO] != READING_BLOCK_COL_HEADING or READING_CONS_COL_HEADING != row[READING_CONS_COL_NO] then
    puts 'The format of the reading Excel file ' + rdngFile + ' does not seems to be correct, check the file and retry!'
    abort
end
sheet.each do |row|
    break if row.join('').empty?
    next if READING_BLOCK_COL_HEADING == row[READING_BLOCK_COL_NO]
    block = row[READING_BLOCK_COL_NO]
    flat = row[READING_FLAT_COL_NO].to_i
    consumed = row[READING_CONS_COL_NO].to_f
    negCount = 0;
    if consumed < 0 then
        puts "ERROR - NEGATIVE CONSUMPTION FOUND FOR #{block}-#{flat}!"
        negCount = negCount + 1
    end
    if consumed > 25 then
        puts "WARNING!! HIGH consumtion, #{consumed} m^3 found for #{block}-#{flat}! Continuing bill generation, but review the readings of this flat!"
    end
    if negCount > 0 then
        if ALLOW_NEGATIVE_READINGS != true then
            puts 'NEGATIVE READINGS FOUND, THERE IS SOMETHING WRONG WITH THE READINGS. CORRECT THE READINGS AND RETRY'
            abort
        else
            puts "Continuing bill generation, but review the readings of these flats!"
        end
    end
    Flat.set_consumption(block, flat, consumed)
    count = count + 1
end
puts "Number of readings entered: #{count}"

#Tests
puts 'Checking data sanity...'
Flat.check_sanity

#Read gas ledger excel file
puts 'Reading file ' + ledgFile
book = Spreadsheet.open ledgFile
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
puts 'Reading file ' + outsdFile
book = Spreadsheet.open outsdFile
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
