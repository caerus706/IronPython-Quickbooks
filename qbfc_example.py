import clr
# Interop is needed to treat it as a .NET object
clr.AddReferenceByPartialName("Interop.QBFC13")

# Import after reference
import Interop.QBFC13
# QBSessionManagerClass needed instead of QBSessionManager
# QBSessionManager is treated as an Abstract and cannot be instantiated
from Interop.QBFC13 import QBSessionManagerClass



# Without Class it considers itself abstract and will not send itself as an object on connection
qb = QBSessionManagerClass()

# Connection will look for Open file if this isn't named
QBFILE = ""

qb.OpenConnection(QBFILE, "Test QBFC Request")

qb.BeginSession("", 0)

# Set quickbooks (US), MajorVer, MinorVer
rqMsg = qb.CreateMsgSetRequest("US", 6, 0)

# Query all Inventory Adjustments
rqMsg.AppendInventoryAdjustmentQueryRq()

# Send request to quickbooks
# Response is a COM object
resMsg = qb.DoRequests(rqMsg)

# end session and close connection
qb.EndSession()
qb.CloseConnection()

QBXML = resMsg

QBXMLMsgRq = QBXML.ResponseList


try:
    InvAdjQueryRes = QBXMLMsgRq.GetAt(0)
except:
    print("Could not get InvAdjustQuery")

# len(InvAdjQueryRes.Detail) did not work, but InvAdjQueryRes.Detail.Count does
for x in range(0, InvAdjQueryRes.Detail.Count):
    try:
        InvAdjRet = InvAdjQueryRes.Detail.GetAt(x)
        # there is always a TxnID
        TxnId = InvAdjRet.TxnID.GetValue()
        try:
            # will throw error if memo doesn't exist        
            memo = InvAdjRet.Memo.GetValue()
        except:
            memo = "no memo"
        print(TxnId, memo)
    except:
        print("Could not read")
        pass
