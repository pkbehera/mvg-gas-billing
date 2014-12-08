require 'constants'
require 'rubygems'
require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

class Flat
    #Static map of all flats
    @@all_flats = {}
    @@flat_count = 0
    #Public Static methods
    def Flat.add_flat(block, unitNo)
        #return if not flat.kind_of? Flat
        key = block + "%04d" % unitNo
        flat = Flat.new(block, unitNo)
        @@all_flats[key] = flat
        @@flat_count += 1
    end
    def Flat.get_flat_count
        @@flat_count
    end
    def Flat.check_sanity
        if Flat.get_flat('A', 1001).to_s != 'A, 1001, true, true' or Flat.get_flat('C', 801).to_s != 'C, 801, true, true' or Flat.get_flat('E', 702).to_s != 'E, 702, true, true' or Flat.get_flat('A', 501).to_s != 'A, 501, false, true' or Flat.get_flat('B', 604).to_s != 'B, 604, false, true' then
            puts 'Something wrong, flat checks not working!'
            abort
        end
    end
    def Flat.set_kyc_subscried(block, unitNo, subscribed, kyc)
        flat = get_flat(block, unitNo)
        flat.subscribed = subscribed
        flat.kyc = kyc
    end
    def Flat.set_subsidy_unac_debit_no_reading_cnt(block, unitNo, subsidy_used, unacc_debit, no_reading_cnt)
        flat = get_flat(block, unitNo)
        flat.subsidy_used = subsidy_used.to_f
        #Subsidy used resets to 0 in the April each year (when billing is done in May)
        if Time.now.month == 5 then
            flat.subsidy_used = 0.0
        end
        flat.unacc_debit = unacc_debit.to_f
        flat.no_reading_cnt = no_reading_cnt.to_i
    end
    def Flat.set_outstanding(block, unitNo, outstanding)
        flat = get_flat(block, unitNo)
        flat.outstanding = outstanding.to_f
    end
    def Flat.set_occ(block, unitNo, occ)
        flat = get_flat(block, unitNo)
        flat.occupied = occ
    end
    def Flat.set_consumption(block, unitNo, consumed)
        flat = get_flat(block, unitNo)
        if flat.reading_avlbl then
            puts "WARNING!! more than one reading for for flat #{block}=#{unitNo}!"
        end
        flat.consumed = consumed
        flat.reading_avlbl = true
    end
    def Flat.process_debits
        row_debit = 0
        row_ledger = 0
        unocc_flats = 0
        unsub_flats = 0
        occ_and_sub_flats = 0
        occ_sub_and_kyc_flats = 0
        occ_sub_and_no_kyc_flats = 0
        flat_debits = 0
        kyc_flats = 0
        non_kyc_flats = 0
        non_flat_debits = 0
        total_billed_amount = 0.0
        total_unadjusted_debit = 0.0
        fined_no_reading = 0
        fined_no_payment = 0
        debit_book = Spreadsheet::Workbook.new
        debit_sheet = debit_book.create_worksheet(:name => OUTPUT_SHEET_NAME)
        debit_sheet.row(row_debit).push(OUTPUT_BLOCK_COL_HEADING).push(OUTPUT_FLAT_COL_HEADING).push(OUTPUT_AMT_COL_HEADING).push(OUTPUT_AC_HEAD_COL_HEADING).push(OUTPUT_COMMENTS_COL_HEADING)
        ledger_book = Spreadsheet::Workbook.new
        ledger_sheet = ledger_book.create_worksheet(:name => OUTPUT_SHEET_NAME)
        ledger_sheet.row(row_ledger).push(OUTPUT_BLOCK_COL_HEADING).push(OUTPUT_FLAT_COL_HEADING).push(OUTPUT_SUBSIDY_COL_HEADING).push(OUTPUT_UNAC_DEBIT_HEADING).push(LEDGER_NO_READING_COUNT_COL_HEADING)

        @@all_flats.sort.each do |key, flat|
            #puts "Processing " + key
            flat.send(:calculate)
            #puts key + "\t" + flat.send(:to_s)
            #Upload excel file should not have a row for zero debits
            if flat.occupied and flat.subscribed then
                occ_and_sub_flats += 1
                if not flat.reading_avlbl then
                    flat_debits += 1
                else
                    non_flat_debits += 1
                end
                if flat.kyc then
                    occ_sub_and_kyc_flats += 1
                else
                    occ_sub_and_no_kyc_flats += 1
                end
            end
            if not flat.occupied then
                unocc_flats += 1
            end
            if not flat.subscribed then
                unsub_flats += 1
            else
                if flat.kyc then
                   kyc_flats += 1
                else
                    non_kyc_flats += 1
                end
            end
            if flat.fined_no_reading then
                fined_no_reading += 1
            end
            if flat.fined_no_payment then
                fined_no_payment += 1
            end
            total_billed_amount += flat.debit_amt
            total_unadjusted_debit += flat.unacc_debit
            #if flat.debit_amt > 0 then
                row_debit += 1
                debit_sheet.row(row_debit).push(flat.block).push(flat.unitNo).push(flat.debit_amt).push(OUTPUT_AC_HEAD_COL_VAL).push(flat.debit_comment)
            #end
            #There must be an entry for each flat in the ledger
            row_ledger += 1
            ledger_sheet.row(row_ledger).push(flat.block).push(flat.unitNo).push(flat.subsidy_used).push(flat.unacc_debit).push(flat.no_reading_cnt)
        end
        if MVG_TOTAL_NUMBER_OF_FLATS != unsub_flats + kyc_flats + non_kyc_flats then
            puts "ERROR - the total number of flats processed #{(unsub_flats + kyc_flats + non_kyc_flats)} does NOT match with #{MVG_TOTAL_NUMBER_OF_FLATS} records in the Occupancy file! There may be something wrong with the input files!"
            abort
        end
        if occ_and_sub_flats != flat_debits + non_flat_debits then
            puts "ERROR - the total number of debits #{(flat_debits + non_flat_debits)} does NOT match with #{occ_and_sub_flats} flats being billed! There may be something wrong with the input files!"
            abort
        end
        if occ_and_sub_flats != occ_sub_and_kyc_flats + occ_sub_and_no_kyc_flats then
            puts "ERROR - the total number of debits #{(occ_sub_and_kyc_flats + occ_sub_and_no_kyc_flats)} does NOT match with  #{occ_and_sub_flats} flats being billed! There may be something wrong with the input files!"
            abort
        end
        time = Time.new
        puts 'Date-Time Run: ' + time.day.to_s + '/' + time.month.to_s + '/' + time.year.to_s + '-' + time.hour.to_s + ':' + time.min.to_s
        puts "------------SETTINGS USED------------"
        puts "Mass conversion rate: #{VOL_MASS_RATIO} kg/m^3"
        puts "Subsidised billing rate: Rs. #{SUBSIDISED_CHARGE_PER_KG} per kg"
        puts "Commercial billing rate: Rs. #{COMMERCIAL_CHARGE_PER_KG} per kg"
        puts "-------------------------------------"
        puts "#{kyc_flats} Flats have completed KYC formalities"
        puts "#{non_kyc_flats} Flats have NOT completed KYC formalities or unsubscribed"
        puts "-------------------------------------"
        puts "#{unsub_flats} Flats were billed ZERO amounts, as they have UNSUBSCRIBED"
        puts "#{unocc_flats} Flats were billed ZERO amounts, as they are UNOCCUPIED"
        puts "#{occ_and_sub_flats} Flats (KYC: #{occ_sub_and_kyc_flats}, NON-KYC: #{occ_sub_and_no_kyc_flats}) were billed NON-ZERO amounts, as they are OCCUPIED and SUBSCRIBED"
        puts "-------------------------------------"
        puts "#{non_flat_debits} Flats were billed as per readings provided"
        puts "#{flat_debits} Flats were billed FIXED amounts as readings NOT provided"
        if fined_no_reading > 0 then
            puts "#{fined_no_reading} Flats were FINED for NOT providing readings for more than #{NO_READING_NO_FINE_CYCLES} consecutive months!"
        end
        if fined_no_payment > 0 then
            puts "#{fined_no_payment} Flats were FINED for NOT paying their previous dues!"
        end
        puts "#{(unsub_flats + kyc_flats + non_kyc_flats)} Flats were processed in total"
        puts "Total billed amount: Rs. #{total_billed_amount}"
        puts "Total un-adjusted debit: Rs. #{total_unadjusted_debit}"
        puts "-------------------------------------"
        #First 3 letters of a Month name
        cur_month = Date::MONTHNAMES[time.month][0..2]
        debit_file = OUTPUT_GAS_DEBIT_FILE_NAME % [cur_month, time.year]
        debit_book.write debit_file
        puts "Debit amounts written to (to be uploaded to OneSolution): #{debit_file}"
        ledger_file = OUTPUT_GAS_LEDGER_FILE_NAME % [cur_month, time.year]
        ledger_book.write ledger_file
        puts "Gas ledger written to (to be used as a input next month): #{ledger_file}"
    end

    attr_accessor :block, :unitNo, :consumed, :occupied, :subscribed, :kyc, :subsidy_used, :unacc_debit, :no_reading_cnt, :fined_no_reading, :outstanding, :fined_no_payment, :reading_avlbl, :debit_amt, :debit_comment, :to_s
    def to_s
        @block.to_s + ', ' + @unitNo.to_s + ', ' + @kyc.to_s + ', ' + @occupied.to_s
        #+ ', ' + @reading_avlbl.to_s + ', ' + @consumed.to_s + ', ' + @total_subsidy_used.to_s + ', ' + @unacc_debit.to_s + ', ' + @debit_amt.to_s + ', ' + @debit_comment
    end
    #Private static methods
    private
    def Flat.get_flat(block, unitNo)
        unit = unitNo.to_s
        if unit.length < 4 then
            unit = '0' + unit
        end
        key = block + unit
        flat = @@all_flats[key]
        if flat.nil? then
            puts 'Could not find flat ' + block + unitNo.to_s + ', there is something wrong!!'
            abort
        end
        return flat
    end

    #Private instance methods
    def initialize(block, unitNo)
        @block = block.to_s
        @unitNo = unitNo.to_i
        #when subscribed = false, the flat is not to be billed
        @subscribed = true
        @kyc = true
        @occupied = true
        @reading_avlbl = false
        #Volumetric unit
        @consumed = 0.0
        @subsidy_used = 0.0
        @unacc_debit = 0.0
        @debit_amt = 0.0
        @debit_comment = ''
        @no_reading_cnt = 0
        @fined_no_reading = false
        @outstanding = 0.0
        @fined_no_payment= false
    end

    def calculate
        @debit_amt = 0
        @debit_comment = DEBIT_COMMENT_UNOCCUPIED
        if not @subscribed then
            @debit_comment = DEBIT_COMMENT_UNSUBSCRIBED
        end
        adjust_amt = 0.0
        if @reading_avlbl and @consumed > 0.0 then
            @occupied = true
            @subscribed = true
        end
        if @occupied && @subscribed then
            if @reading_avlbl then
                @no_reading_cnt = 0
                used_kgs = consumed * VOL_MASS_RATIO
                rem_subsidy_kgs = TOTAL_SUBSIDY_PER_YR_KGS - @subsidy_used
                if rem_subsidy_kgs < 0.0 then
                    rem_subsidy_kgs = 0.0
                end
                subsidized_kgs = used_kgs
                non_subsidized_kgs = 0.0
                @debit_comment = DEBIT_COMMENT_SUBSIDISED
                if @kyc then
                    if rem_subsidy_kgs < used_kgs then
                        subsidized_kgs = rem_subsidy_kgs
                        non_subsidized_kgs = used_kgs - subsidized_kgs
                        @debit_comment = DEBIT_COMMENT_PART_SUBSIDISED
                    end
                    if rem_subsidy_kgs == 0.0 then
                        @debit_comment = DEBIT_COMMENT_COMMERCIAL_KYC
                    end
                    @subsidy_used += subsidized_kgs
                else
                    subsidized_kgs = 0.0
                    #This is to provide pro-rated subsidy to people who do their KYC in the middle of the financial year
                    @subsidy_used += SUBSIDY_PER_MONTH
                    non_subsidized_kgs = used_kgs
                    @debit_comment = DEBIT_COMMENT_COMMERCIAL_NO_KYC
                end
                @debit_amt = subsidized_kgs * SUBSIDISED_CHARGE_PER_KG + non_subsidized_kgs * COMMERCIAL_CHARGE_PER_KG
                sub_com = (subsidized_kgs > 0 ? "#{(subsidized_kgs*1000).round/1000.0} kg * #{SUBSIDISED_CHARGE_PER_KG}" : '')
                com_com = (non_subsidized_kgs > 0 ? "#{(non_subsidized_kgs*1000).round/1000.0} kg * #{COMMERCIAL_CHARGE_PER_KG}" : '')
                comment = sub_com
                if sub_com != '' && com_com != '' then
                    comment = sub_com + ' + ' + com_com
                elsif sub_com == '' then
                    comment = com_com
                end
                if comment != '' then
                    @debit_comment += ' [' + comment + ']'
                end
                if @unacc_debit > 0.0 then
                    debit = @debit_amt - @unacc_debit
                    @unacc_debit = 0.0
                    if debit < 0.0 then
                        @unacc_debit = debit * -1.0
                        debit = 0.0
                    end
                    adjust_amt = @debit_amt - debit
                    if adjust_amt > 0 then
                        adjust_amt = (adjust_amt * 100).round / 100.0
                        @debit_amt = (@debit_amt * 100).round / 100.0
                        @unacc_debit = (@unacc_debit * 100).round / 100.0
                        @debit_comment = @debit_comment.chop + "=#{@debit_amt}, adjusted #{adjust_amt} against previous debits, remaining unadjusted debit #{@unacc_debit}]"
                    end
                    @debit_amt = debit
                end
            else
                @no_reading_cnt += 1
                if @kyc then
                    @debit_amt = NO_READING_SUBSIDISED_DEBIT
                    @debit_comment = DEBIT_COMMENT_NO_READING_KYC
                else
                    @debit_amt = NO_READING_COMMERCIAL_DEBIT
                    @debit_comment = DEBIT_COMMENT_NO_READING_NO_KYC
                    #Provide pro-rated subsidy to people who complete KYC in the middle of a financial year
                    @subsidy_used += SUBSIDY_PER_MONTH
                end
                @unacc_debit = @unacc_debit + @debit_amt
                #Fine for not providing readings
                if APPLY_FINE_NO_READING_DEFAULTS and @no_reading_cnt > NO_READING_NO_FINE_CYCLES then
                    fine_amt = @no_reading_cnt == NO_READING_NO_FINE_CYCLES + 1 ? NO_READING_FINE_AMT*(NO_READING_NO_FINE_CYCLES + 1) : NO_READING_FINE_AMT
                    @debit_amt += fine_amt
                    @debit_comment += " #{fine_amt} fine for not providing readings for more than #{NO_READING_NO_FINE_CYCLES} consecutive months"
                    @fined_no_reading = true
                end
            end
        else
            #Provide pro-rated subsidy to people who complete KYC or subscribe in the middle of a financial year
            if @subsidy_used == 0.0 then
                month_num = Time.now.month
                if month_num < 5 then
                    month_num += 12
                end
                @subsidy_used = SUBSIDY_PER_MONTH*(month_num - 4)
            else
                @subsidy_used += SUBSIDY_PER_MONTH
            end
        end
        #Interest for not making payments in time
        if APPLY_INTEREST_NO_PAYMENTS and @outstanding > PENAL_INTEREST_THRESHOLD then
            interest = @outstanding * PENAL_INTEREST_NO_PAYMENTS / 100.00
            if interest < MIN_PENAL_INTEREST_NO_PAYMENTS then
                interest = MIN_PENAL_INTEREST_NO_PAYMENTS
            end
            @debit_amt += interest
            @debit_comment += " #{interest} interest for not paying previous dues:#{@outstanding}"
            @fined_no_payment = true
        end
        if ALLOW_REFUNDS and @unacc_debit > 0 then
            if not @subscribed or not @occupied or @reading_avlbl then
                refund = @unacc_debit
                @unacc_debit = 0.0
                @debit_amt += -1 * refund
                @debit_comment = @debit_comment.split("remaining unadjusted debit ")[0] + " reversed #{refund}, remaining unadjusted debit #{@unacc_debit}]"
           end
        end
        #Now Round off
        #Three decimal places
        @subsidy_used = (@subsidy_used * 1000).round / 1000.0
        #Two decimal places
        @debit_amt = (@debit_amt * 100).round / 100.0
        @unacc_debit = (@unacc_debit * 100).round / 100.0
        if @kyc then
            @debit_comment += " Subsidy used:#{@subsidy_used} out of #{TOTAL_SUBSIDY_PER_YR_KGS} KG"
        end
    end
end
