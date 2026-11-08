# GetNameList

## Syntax

SapObject.SapModel.CableObj.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of cable object names retrieved by the program.

MyName

This is a one-dimensional array of cable object names. The MyName array is created as a dynamic, zero-based, array by the APIuser:



```
   Dim MyName() as String

 
```



The array is dimensioned to (NumberNames – 1) inside the Sap2000 program, filled with the names, and returned to the APIuser.

## Remarks

This function retrieves the names of all defined cable objects.

The function returns zero if the names are successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetCableObjectNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberNames As Long

      Dim MyName() As String

      Dim Name1 As String

      Dim Name2 As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name1)

      ret = SapModel.CableObj.AddByPoint("5", "10", Name2)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name1, 7, 1, 0, 0, 24)

      ret = SapModel.CableObj.SetCableData(Name2, 7, 1, 0, 0, 24)



   'get cable object names

      ret = SapModel.CableObj.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
