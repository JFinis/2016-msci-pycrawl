import urllib.request
import sys, bs4
import json
import re
from lxml.html import soupparser
from lxml.etree import tostring
import locale
import datetime



locale.setlocale( locale.LC_ALL, 'en_US.UTF-8' )


# Capitalization options
# <select name="templateForm:_id80" size="1" class="aoc-ComboBox paragraphTextFont ">   
#<option value="111">A-Series</option>    
#<option value="77">All Cap (Large+Mid+Small+Micro Cap)</option>    
#<option value="108">All Market</option>    
#<option value="41">IMI (Large+Mid+Small Cap)</option>    
#<option value="37">Large Cap</option>    
#<option value="76">Micro Cap</option>    
#<option value="38">Mid Cap</option>    
#<option value="119">Provisional IMI</option>    
#<option value="99">Provisional Small Cap</option>    
#<option value="29">Provisional Standard</option>    
#<option value="40">SMID (Small+Mid Cap)</option>    
#<option value="79">Small + Micro Cap</option>    
#<option value="39">Small Cap</option>    
#<option value="36" selected="selected">Standard (Large+Mid Cap)</option>
requestedCapitalizations = ["36","39","38","41","37","76"]
capNames = { "36":"LM","39":"S","38":"M","41":"LMS","37":"L","76":"m"}

# Style options
#<select name="templateForm:_id96" size="1" class="aoc-ComboBox paragraphTextFont ">
#<option value="G">Growth</option>
#<option value="C" selected="selected">None</option>
#<option value="V">Value</option></select>\
requestedStyles = ["C","V","G"]
styleNames = {"C":"normal","V":"value","G":"growth"}

# Market options
#<select name="templateForm:_id56" size="1" class="aoc-ComboBox paragraphTextFont ">
#<option value="1896">All Country (DM+EM)</option>
#<option value="2809">China Markets</option>
#<option value="1897" selected="selected">Developed Markets (DM)</option>
#<option value="1898">Emerging Markets (EM)</option>
#<option value="2115">Frontier Markets (FM)</option>
#<option value="1899">GCC and Arabian Markets</option></select>
requestedMarkets = ["1897","1898","2115","1896"]
marketNames = {"1897":"DM","1898":"EM","2115":"FM","1896":"AC"}


class UnanticipatedContentException(Exception):
    
    def __init__(self,message):
        super(UnanticipatedContentException, self).__init__(message)


def dumpPretty(response):
    root = soupparser.fromstring(response)
    sys.stdout.write(tostring(root, pretty_print=True).decode('utf-8'))

def checkElement(check,message):
    if not check:
        raise UnanticipatedContentException(message)
    
def assertOne(list,row):
    if len(list) != 1:
        raise UnanticipatedContentException(tostring(row))
    return list
    
def parseMsciDataRow(row):
    indexName = assertOne(row.xpath("./td[1]/a/text()"),row)[0]
    indexCode = assertOne(row.xpath("./td[2]/text()"),row)[0]
    dayValueStr = assertOne(row.xpath("./td[3]/text()"),row)[0]
    dayValue =  locale.atof(dayValueStr)
    return (indexName,indexCode,dayValue)

def parseMsciResponse(requestedDate,responseStr,outFile,market,cap,style,resultNum):
    # Parse to XML tree
    root = soupparser.fromstring(responseStr)
    
    # Find names
    marketName = marketNames[market]
    capName = capNames[cap]
    styleName = styleNames[style]
    
    # First, get thee header text such as "All Country Standard  (Net) as of Aug 05, 2016"
    headerText = root.xpath('//table/tr/td/span[@class="paragraphTextFont "]/text()')
    # Sanity check: one header with the right text
    checkElement(len(headerText)==1 and "as of" in headerText[0],headerText)
    
    # Extract header data, make sure it is the data we requested
    match = re.compile("([A-Za-z ]+)\s+\((.*)\) as of (.*)").match(headerText[0])
    checkElement (match != None,headerText[0])
    indexGroup=match.group(1)
    indexLevel=match.group(2)
    asOfDate=match.group(3)
    checkElement(asOfDate==requestedDate,asOfDate);
    checkElement(indexLevel=="Net",indexLevel);
    
    # Now extract the data rows
    dataRows = root.xpath('//tbody[@id="templateForm:tableResult0:tbody_element"]/tr')
    
    # Check for an empty table (no data for this day)
    if dataRows[0].xpath("./td[1]//text()")==[]:
        if len(dataRows)!=1 :
            raise Exception()
        resultStr="No data for\t" + requestDate + "\t" + marketName + "\t" + capName + "\t" + styleName + "\t" + "("+ str(resultNum) + ")"
        outFile.write("#" + resultStr + "\n")
        print(resultStr)
        return
    

    # Write header comment to file    
    resultStr="Acquired " + str(len(dataRows)) + " rows for\t" + requestDate + "\t" + marketName + "\t" + capName + "\t" + styleName + "("+ str(resultNum) + ")"
    outFile.write("#" + resultStr + "\n")
    
    # Write data rows to file
    for row in dataRows:
        (indexName,indexCode,dayValue)=parseMsciDataRow(row)
        outFile.write(str(resultNum) + "\t" + marketName + "\t" + capName + "\t" + styleName + "\t" + indexLevel + "\t" + asOfDate + "\t" + indexName + "\t" + indexCode + "\t" + str(dayValue) + "\n")

    # Flush file, print summary
    outFile.flush()
    print(resultStr)
    

def requestMsciPage(requestedDate,requestedMarket,requestedCapitalization,requestedStyle):
    params = {"templateForm:_id51":"08/06/2016",
              "templateForm:_id56":requestedMarket,
              "templateForm:_id66":requestedDate,
              "templateForm:selectOneMenuCategoryCurrency":"119", # 15 = USD 119 = EUR
              "templateForm:_id80":requestedCapitalization,
              "templateForm:selectOneMenuCategoryLevel":"41",
              "templateForm:_id96":requestedStyle,
              "templateForm:selectOneMenuCategoryFamily":"C",
              "templateForm:_id106":"Search",
              "templateForm_SUBMIT":"1",
              "templateForm:_link_hidden_":"",
              "templateForm:_idcl":"",
              }
    paramsEncoded = urllib.parse.urlencode(params).encode('utf-8')
    req = urllib.request.Request("https://app2.msci.com/webapp/indexperf/pages/IEIPerformanceRegional.jsf",
                                 data=paramsEncoded,headers={"Cookie":"JSESSIONID=4A64F7301E10ADB2713192E15EC02C2E; _ga=GA1.2.1988533469.1475785413; NSC_JO4gtwr2cl4jwkmd3nrmb3ck0yhebb3=ffffffff09371ea345525d5f4f58455e445a4a422970"})
    return urllib.request.urlopen(req).read()

if __name__ == '__main__':

    # Open file, determine start date
    outFile = open('data.csv', 'a')
    date_1 = datetime.datetime.strptime("11/01/2011", "%m/%d/%Y")

    # Do the scraping
    index=0
    for cap in requestedCapitalizations:
        for style in requestedStyles:
            for market in requestedMarkets:
                for i in range(1796):
                    # Assemble date
                    end_date = date_1 + datetime.timedelta(days=i)
                    requestDate = end_date.strftime("%b %d, %Y")
                
                    # Do the request
                    response=requestMsciPage(requestDate,market,cap,style)
                    #dumpPretty(response)
                    parseMsciResponse(requestDate,response,outFile,market,cap,style,index)
                    index=index+1
        
    outFile.close()
        