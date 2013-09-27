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

# Here you can define the currency used with your account (e.g. EUR, SEK)
MY_CURRENCY = "SGD"


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


def convertFile(f):
    """
    Creates new OFX file, then maps transactions from original CSV (f) to
    OFX's version of XML.
    
    @param f: A file handle (f) for the original CSV transactions list.
    @type f: File
    """

    # Open/create the .ofx file
    try:
        outFile = open(arg.split('.')[0] + ".ofx", "w")
    except IOError:
        print("Output file couldn't be created. Program will now exit")
        sys.exit(2)

    # Create csv reader and read in account number
    csvReader = csv.reader(f, dialect=csv.excel_tab)
    acctDetailsFor_line1 = csvReader.next() # TODO: what is this line for?
    acctDetailsFor_line2 = csvReader.next()[0] # TODO: verify that this is the line with the acct number
    acctNumber = "".join(re.findall(r"\d{3}-\d{6}-\d{3}", acctDetailsFor_line2)) # Converted to string

    # Get info from file name about dates (time is not given, so we add 12:00
    # as arbitrary time)

    # TODO: parsing this with a REGEX would be less fragile
    # TODO: make the parsing of the dates automatic
    try:
        dateStart = f.name.split('_')[2] + "120000"
        dateEnd = f.name.split('_')[3].split('.')[0] + "120000"
    except IndexError:
        print "Unable to automatically retrieve the start/end dates for your file."
        print "Please enter the start/end dates in the following format: YYYYMMDD (8 digits)."
        dateStart = raw_input("Please enter a start date: ") + "12000"
        dateEnd = raw_input("Please enter an end date: ") + "12000"

    # Bypasses unneeded lines
    while csvReader.line_num < 7:
        unnecessary_line = csvReader.next()

    # Creates string from file's time stamp
    timeStamp = time.strftime(
        "%Y%m%d%H%M%S", time.localtime((os.path.getctime(f.name))))
    # Version with timezone: timeStamp = time.strftime("%Y%m%d%H%M%S" + "[+2:%Z]", time.localtime())

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
                        <DTSERVER>''' + timeStamp +  '''</DTSERVER>
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
                    <DTSTART>''')

    # Read lines from csvReader and add them as transactions
    for line in csvReader:
        if line != []:
            # Unpacks the line into variables (including one null cell). Reformats the date.
            entryDate, valueDate, paymentDate, amount, name, account, bic,\
            transaction, refNum, refOrigNum, message, cardNum,\
            receipt, nullCell = line

            entryDate = entryDate.split('.')[2] + entryDate.split(
                '.')[1] + entryDate.split('.')[0] + "120000"

            # Quick and dirty trans type (needs a function table)
            outFile.write(
                '''<STMTTRN>
                        <TRNTYPE>''' + getTransType(transaction, amount) + '''</TRNTYPE>
                        <DTPOSTED>''' + entryDate + '''</DTPOSTED>
                        <TRNAMT>''' + amount + '''</TRNAMT>
                        <FITID>''' + refNum + '''</FITID>
                        <NAME>''' + name + '''</NAME>
                        <MEMO>''' + message + '''</MEMO>
                    </STMTTRN>
                    ''')

    # Write OFX footer
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
            f_in = open(arg, "rb")
            print("Opening %s" % arg)
            convertFile(f_in)
        except IOError:
            print("Error: file %s couldn't be opened" % arg)
        else:
            f_in.close()
            print("%s is closed" % arg)
