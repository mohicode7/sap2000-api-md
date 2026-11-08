# ShowTablesInExcel

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.ShowTablesInExcel

## VB6 Procedure

Function ShowTablesInExcel(ByRef TableKeyList() As String,
WindowHandle As Integer) As Integer

## Parameters

TableKeyList

A zero-based array of the table keys for the tables
to be included in the output.

WindowHandle

If a valid window handle is
entered here Excel will be located in front of that window.  

## Remarks

Exports the specified tables to Excel.

Excel must be present on the computer for this function
to work. If there is nothing to be shown in the table then no data is
returned and Excel is not displayed.

## VBA Example



```
Sub Example()
 
    Dim myHelper As cHelper
    Dim SapObject As cOAPI
    Dim SapModel As cSapModel
    Dim ret As Long
    ret = -1
   
    'create SapObject
    Set myHelper = New Helper
    Set SapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
   
    'start SAP2000 application
    ret = SapObject.ApplicationStart()
   
    'create SapModel object
    Set SapModel = SapObject.SapModel
   
    'Initialize model
    ret = SapModel.InitializeNewModel()
 
    'Create 2D Frame Model
    ret = SapModel.File.New2DFrame(e2DFrameType_PortalFrame, 
 2, 10, 3, 20)
   
    'Get available tables
    Dim NumberTables As Long
    Dim TableKey() As String
    Dim TableName() As String
    Dim ImportType() As Long
    Dim IsEmpty() As Boolean
    ret = SapModel.DatabaseTables.GetAvailableTables(NumberTables, 
 TableKey, TableName, ImportType)
   
    Dim TableKeyList() As String
    ReDim TableKeyList(2)
    TableKeyList(0) = TableKey(5) 
 ' "Connectivity - Frame"
    TableKeyList(1) = TableKey(21) 
 ' "Grid Lines"
    TableKeyList(2) = TableKey(31) 
 ' "Material Properties 03a - Steel Data"
    Dim WindowHandle As Long
    WindowHandle = 1
   
    'View some tables in Excel
    ret = SapModel.DatabaseTables.ShowTablesInExcel(TableKeyList, 
 WindowHandle)
   
    'close Sap2000
    SapObject.ApplicationExit (False)
    Set SapModel = Nothing
    Set SapObject = Nothing
End Sub
 
```



## Release Notes

Initial release in version 23.0.0.

Added VBA Example in 24.2.0.

## See Also
