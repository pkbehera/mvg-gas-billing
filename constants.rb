#Constants
#These may change every month
VOL_MASS_RATIO = 2.26
SUBSIDISED_CHARGE_PER_KG = 34.47
NON_SUBSIDISED_CHARGE_PER_KG = 48.30
#Used till Apr 2013 billing (done in May'13)
#COMMERCIAL_CHARGE_PER_KG = 74.36
#Used till Nov 2013 billing (done in Dec'13)
#COMMERCIAL_CHARGE_PER_KG = 71.69
#Used till Dec 2013 billing (done in Jan'14)
#COMMERCIAL_CHARGE_PER_KG = 91.28
COMMERCIAL_CHARGE_PER_KG = 83.90
#Used till Dec 2013 billing (done in Jan'14)
#TOTAL_SUBSIDY_PER_YR_KGS = 14.10 * 9
TOTAL_SUBSIDY_PER_YR_KGS = 14.10 * 12
SUBSIDY_PER_MONTH = TOTAL_SUBSIDY_PER_YR_KGS / 12.00
NO_READING_KYC_DEBIT  = 300.00
NO_READING_NO_KYC_DEBIT  = 600.00
NO_READING_NO_FINE_CYCLES = 3
NO_READING_FINE_AMT = 50.00
APPLY_FINE_NO_READING_DEFAULTS = true
APPLY_INTEREST_NO_PAYMENTS = false
#Percent per month
PENAL_INTEREST_NO_PAYMENTS = 2.00
MIN_PENAL_INTEREST_NO_PAYMENTS = 50.00
PENAL_INTEREST_THRESHOLD = 0.00
ALLOW_REFUNDS = true
ALLOW_NEGATIVE_READINGS = false
HIGH_CONS_WARN = 40

#MVG FLATS
MVG_ALL_BLOCKS = ['A', 'B', 'C', 'E']
MVG_FLRS_IN_EACH_BLK = [1..11, 1..10, 1..11, 1..11]
MVG_FLTS_IN_EACH_FLR = [1..6, 1..6, 1..6, 1..5]
MVG_FIRE_REFUSES = ['A803', 'B805', 'C803', 'E804']
#A1105 & A1106 have a common kitchen
MVG_IGNORED_FLATS = ['A1105']

#Used for sanity check of input Excel files
#Total number of flats, excluding fire refuses and ignored flats
MVG_TOTAL_NUMBER_OF_FLATS = 242

#Debit Comments
DEBIT_COMMENT_UNOCCUPIED = 'No debit - Flat Unoccupied'
DEBIT_COMMENT_UNSUBSCRIBED = 'No debit - Flat Un-Subscribed'
DEBIT_COMMENT_NO_READING_KYC = 'No reading - flat debit'
DEBIT_COMMENT_NO_READING_NO_KYC = 'No reading - flat debit, KYC NOT done'
DEBIT_COMMENT_SUBSIDISED = 'Subsidised rate'
DEBIT_COMMENT_PART_SUBSIDISED = 'Partly non-subsidised rate, consumed beyond subsidy quota'
DEBIT_COMMENT_NON_SUBSIDISED_KYC = 'Non-subsidised rate, consumed beyond subsidy quota'
DEBIT_COMMENT_NON_SUBSIDISED_NO_KYC = 'Non-subsidised rate, KYC NOT done'
DEBIT_COMMENT_COMMERCIAL_UNSUBSCRIBED = 'Commercial rate, flat unsubscribed but consuming gas!'

#Constants for file containing KYC Values
KYC_BLOCK_COL_NO = 1 #Column B
KYC_FLAT_COL_NO = 2  #Column C
KYC_KYC_COL_NO = 3   #Column D
OCC_OCC_COL_NO = 4   #Column E
KYC_KYC_COL_HEADING = 'KYC Status'
KYC_KYC_OK_STRING_DOWNCASE = 'yes'
KYC_BLOCK_COL_HEADING = 'Block'
KYC_FLAT_COL_HEADING = 'Unit No.'
OCC_OCC_COL_HEADING = 'Occupancy Status'
OCC_UNOCC_STRING_DOWNCASE = 'unoccupied'
OCC_BLOCK_COL_HEADING = KYC_BLOCK_COL_HEADING
OCC_FLAT_COL_HEADING = KYC_FLAT_COL_HEADING

#Constants for file containing latest readings
#These would change if the structure of the Excel sheet from onesolution with Gas readings changes
READING_BLOCK_COL_NO = 1 #Column B
READING_FLAT_COL_NO = 2  #Column C
READING_CONS_COL_NO = 7  #Column H
READING_BLOCK_COL_HEADING = KYC_BLOCK_COL_HEADING
READING_FLAT_COL_HEADING = KYC_FLAT_COL_HEADING
READING_CONS_COL_HEADING = 'Units Consumed'

#Constatnt for file containing outstandings, TODO - update these
OUTSTAND_BLOCK_FLAT_COL_NO = 1 #Column B
OUTSTAND_BAL_COL_NO = 9  #Column J
OUTSTAND_BLOCK_FLAT_COL_HEADING = 'Unit No'
OUTSTAND_BAL_COL_HEADING = 'Domestic Gas'

#Constants for Gas ledger
LEDGER_BLOCK_COL_NO = 0
LEDGER_FLAT_COL_NO = 1
LEDGER_SUB_USED_COL_NO = 2
LEDGER_UNAC_DEBIT_COL_NO = 3
LEDGER_NO_READING_COUNT_COL_NO = 4
LEDGER_BLOCK_COL_HEADING = KYC_BLOCK_COL_HEADING
LEDGER_FLAT_COL_HEADING = 'flatno'
LEDGER_UNAC_DEBIT_COL_HEADING = 'Unadjusted Debit'
LEDGER_NO_READING_COUNT_COL_HEADING = 'No Reading Count'

#These would change if the structure of the Excel sheet uploaded to onesolution changes
OUTPUT_SHEET_NAME = 'Sheet1'
OUTPUT_BLOCK_COL_HEADING = LEDGER_BLOCK_COL_HEADING
OUTPUT_FLAT_COL_HEADING = LEDGER_FLAT_COL_HEADING
OUTPUT_AMT_COL_HEADING = 'Amount'
OUTPUT_AC_HEAD_COL_HEADING = 'achead'
OUTPUT_COMMENTS_COL_HEADING = 'Comments'
OUTPUT_SUBSIDY_COL_HEADING = 'Subsidy Used'
OUTPUT_UNAC_DEBIT_HEADING = LEDGER_UNAC_DEBIT_COL_HEADING
OUTPUT_AC_HEAD_COL_VAL = 'Domestic Gas'
OUTPUT_GAS_DEBIT_FILE_NAME = "Gas_Debits_Upload_%s_%d.xls"
OUTPUT_GAS_LEDGER_FILE_NAME = "Gas_Ledger_%s_%d.xls"
