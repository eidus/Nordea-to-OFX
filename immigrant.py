#!/usr/bin/env python

"""
Immigrant is a small script forked from Nordea to OFX by jgoney
(https://github.com/jgoney/Nordea-to-OFX)

It converts OCBC transaction lists (that are in some silly CSV format) to OFX,
for use with modern financial management software

"Immigrant" is a play on the name, Overseas Chinese Bank. 

Immigrant: converts OCBC transaction lists (CSV) to OFX for use with
financial management software.
"""

import csv
import os
import sys
import time
import re
from datetime import *
import time

# Here you can define the currency used with your account (e.g. EUR, SEK)
MY_CURRENCY = "SGD"

# ============= CONSTANTS ================
# Row number where the Field Names (e.g. transaction date, value date etc) are at (necessary for CSV Dictreader to know!)
FIELDNAMES_LINE_NUMBER = 6  


def getTransType(trans, amt):
    """
    Converts a transaction description (e.g. "Deposit") to an OFX
    standardized transaction (e.g. "DEP").
    
    @param trans: A textual description of the transaction (e.g. "Deposit")
    @type trans: String
    @param amt: The amount of a transaction, used to determine CREDIT or DEBIT.
    @type amt: String
    
    @return: The standardized transaction type
    @rtype: String
    """
    if trans == "ATM withdr/Otto." or trans == "Debit cash withdrawal":
        return "ATM"
    elif trans == "Deposit":
        return "DEP"
    elif trans == "Deposit interest":
        return "INT"
    elif trans == "Direct debit":
        return "DIRECTDEBIT"
    elif trans == "e-invoice" or trans == "e-payment":
        return "PAYMENT"
    elif trans == "ePiggy savings transfer" or trans == "Own transfer":
        return "XFER"
    elif trans == "Service fee VAT 0%":
        return "FEE"
    else:
        if amt[0] == '-':
            return "DEBIT"
        else:
            return "CREDIT"

def getTransAmount(deposits, withdrawals):
    if withdrawals: 
        return "-" + withdrawals
    else:
        return deposits

def convertFile(f):
    """
    Creates new OFX file, then maps transactions from original CSV (f) to
    OFX's version of XML.
    
    @param f: A file handle (f) for the original CSV transactions list.
    @type f: File
    """

    # Open/create the .ofx file
    try:
        outFile = open("output.ofx", "w")
    except IOError:
        print("Output file couldn't be created. Program will now exit")
        sys.exit(2)

    # Reads in the account number
    csvReader = csv.reader(f, dialect=csv.excel_tab)
    acctDetails = csvReader.next()[0] # TODO: verify that this is the line with the acct number
    acctNumber = "".join(re.findall(r"\d{3}-\d{6}-\d{3}", acctDetails)) # Converted to string
    print "Account Number: " + acctNumber

    # SKIPS THE FIRST FEW LINES
    # see http://stackoverflow.com/questions/7588426/how-to-skip-pre-header-lines-with-csv-dictreader
    f.seek(0)
    for i in range(0,FIELDNAMES_LINE_NUMBER-1):    # TODO: pass "5" in as a paramter (in constants) in case OCBC changes its format
        next(f)
    csv_dictReader = csv.DictReader(f)

    # Reads csv transactions into list
    transactionEntries = []
    for line in csv_dictReader:
        print line
        transactionEntries.append(line)

    # Simple test
    numEntries = 0

    # Gets the start and end dates
    startDate = datetime.now()
    endDate = datetime.min
    for entry in transactionEntries:
        if entry['Transaction date']:
            entryDate = datetime.strptime(entry['Transaction date'], "%d/%m/%Y")
            if entryDate < startDate: startDate = entryDate
            if entryDate > endDate: endDate = entryDate
            numEntries += 1
    startDateString = startDate.strftime('%Y%m%d')   # Arbitrary time of 000000 assigned for each start date/end date
    endDateString =  endDate.strftime('%Y%m%d')

    # Creates string from file's time stamp
    timeStamp = time.strftime(
        "%Y%m%d%H%M%S", time.localtime((os.path.getctime(f.name))))

    # Write header to file (includes timestamp)
    outFile.write(
        '''<?xml version="1.0" encoding="ANSI" standalone="no"?>
<?OFX OFXHEADER="200" VERSION="200" SECURITY="NONE" OLDFILEUID="NONE" NEWFILEUID="NONE"?>
<OFX>
        <SIGNONMSGSRSV1>
                <SONRS>
                        <STATUS>
                                <CODE>0</CODE>
                                <SEVERITY>INFO</SEVERITY>
                        </STATUS>
                        <DTSERVER>''' + timeStamp + '''</DTSERVER>
                        <LANGUAGE>ENG</LANGUAGE>
                </SONRS>
        </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>0</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <STMTRS>
                <CURDEF>''' + MY_CURRENCY + '''</CURDEF>
                <BANKACCTFROM>
                    <BANKID>Nordea</BANKID>
                    <ACCTID>''' + acctNumber + '''</ACCTID>
                    <ACCTTYPE>CHECKING</ACCTTYPE>
                </BANKACCTFROM>
                <BANKTRANLIST>
                    <DTSTART>''' + startDateString + '''</DTSTART>
                    <DTEND>''' + endDateString + '''</DTEND>
                    ''')

    numTransactions = 0
    while len(transactionEntries):
        numTransactions += 1
        currentTransaction = transactionEntries.pop(0)
        
        while len(transactionEntries):
            if transactionEntries[0]['Transaction date']: break
            additionalDescription = transactionEntries.pop(0)['Description']
            currentTransaction['Description'] += " " + additionalDescription

        # Just adds a default time of 000000 to each of the transactions
        dateVector = currentTransaction['Transaction date'].split('/')
        print dateVector
        if len(dateVector[1]) < 2: dateVector[1] = '0' + dateVector[1]
        if len(dateVector[0]) < 2: dateVector[0] = '0' + dateVector[0]
        entryDate = dateVector[2] + dateVector[1] + dateVector[0]
        print entryDate
        entryAmount = getTransAmount(currentTransaction['Deposits (SGD)'], currentTransaction['Withdrawals (SGD)'])
        # print currentTransaction

        # Quick and dirty trans type (needs a function table)
        outFile.write(
                '''<STMTTRN>
                        <TRNTYPE>''' + "getTransType(transaction, amount)" + '''</TRNTYPE>
                        <DTPOSTED>''' + entryDate + '''</DTPOSTED>
                        <TRNAMT>''' + entryAmount + '''</TRNAMT>
                        <FITID>''' + 'refNum' + '''</FITID>
                        <NAME>''''''</NAME>
                        <MEMO>''' + currentTransaction['Description'] + '''</MEMO>
                    </STMTTRN>
                    ''')
        
    print "Num Transactions: " + str(numTransactions)
    print "Num Entries: " + str(numEntries)
    
    outFile.write(
        '''</BANKTRANLIST>
                        </STMTRS>
                </STMTTRNRS>
        </BANKMSGSRSV1>
</OFX>''')

    outFile.close()

if __name__ == '__main__':
    # Check that the args are valid
    if len(sys.argv) < 2:
        print("Error: no filenames were given.\nUsage: %s [one or more file names]" % sys.argv[0])
        sys.exit(1)

    # Open the files and put the handles in a list
    for arg in (sys.argv[1:]):
        try:
            f_in = open(arg, "rU")
            print("Opening %s" % arg)
            convertFile(f_in)
        except IOError:
            print("Error: file %s couldn't be opened" % arg)
        else:
            f_in.close()
            print("%s is closed" % arg)
