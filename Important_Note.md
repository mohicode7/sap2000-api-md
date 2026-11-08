# Important Note

This documentation applies to both SAP2000 and CSiBridge.

All examples in this documentation refer to SAP2000. However, users of CSiBridge should refer to CSiBridge. For example,

SAP2000 users can create an instance of the SapObject in VBA as follows



```
Dim myHelper As SAP2000v1.cHelper
Dim mySapObject As SAP2000v1.cOAPI
Set myHelper= New SAP2000v1.Helper
Set mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
```



whereas CSiBridge users would create an instance of the SapObject as follows



```
Dim myHelper As CSiBridge1.cHelper
Dim mySapObject As CSiBridge1.cOAPI
Set myHelper= New CSiBridge1.Helper
Set mySapObject = myHelper.CreateObjectProgID("CSI.CSiBridge.API.SapObject")
```



Other references to SAP2000/SAP2000.exe/SAP2000v1.dll/SAP2000v1.tlb in the documentation should be interpreted as CSiBridge/CSiBridge.exe/CSiBridge1.dll/CSiBridge1.tlb for users of CSiBridge.

Similarly, references to the .sdb extension in SAP2000 should be interpreted as the .bdb extension for users of CSiBridge.
