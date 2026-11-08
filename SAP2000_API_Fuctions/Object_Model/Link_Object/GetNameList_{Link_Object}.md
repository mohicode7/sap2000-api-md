# GetNameList

## Syntax

SapObject.SapModel.LinkObj.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of link object names retrieved by the program.

MyName

This is a one-dimensional array of link object names. The MyName array is created as a dynamic, zero-based, array by the APIuser:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames – 1) inside the SAP2000 program, filled with the names, and returned to the APIuser.

## Remarks

This function retrieves the names of all defined link objects.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetLinkObjectNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberNames As Long

      Dim MyName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("3", "", Name, True)

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'get link object names

      ret = SapModel.LinkObj.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
