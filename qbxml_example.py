import xml.etree.ElementTree
import clr
clr.AddReferenceByPartialName("Interop.QBXMLRP2")

# IronPython needs to add Class to request processor to instantiate requests
from  Interop.QBXMLRP2 import RequestProcessor2Class

# Will look for open file if no filename provided
QBFILE = ''

sessionManager = RequestProcessor2Class()
sessionManager.OpenConnection(QBFILE, 'Test qbXML Request')
ticket = sessionManager.BeginSession("", 0)

# Send query and receive response
qbxml_query = """
<?qbxml version="6.0"?>
<QBXML>
   <QBXMLMsgsRq onError="stopOnError"> 
      <InventoryAdjustmentQueryRq metaData="MetaDataAndResponseData">
      </InventoryAdjustmentQueryRq>
   </QBXMLMsgsRq> 
</QBXML>
"""
response_string = sessionManager.ProcessRequest(ticket, qbxml_query)

# Disconnect from Quickbooks
sessionManager.EndSession(ticket)     # Close the company file
sessionManager.CloseConnection()      # Close the connection

# print string is there to study response
# print(response_string)

# Parse the response into an Element Tree and peel away the layers of response
QBXML = xml.etree.ElementTree.fromstring(response_string)
QBXMLMsgsRs = QBXML.find('QBXMLMsgsRs')
InventoryAdjustmentQueryRs = QBXMLMsgsRs.getiterator("InventoryAdjustmentRet")
for InvAdjRet in InventoryAdjustmentQueryRs:
    txnid = InvAdjRet.find('TxnID').text
    try:
        # in case memo is not entered on adjustment
        memo = InvAdjRet.find('Memo').text
    except:
        memo = "No Memo"

    print(txnid, memo)
