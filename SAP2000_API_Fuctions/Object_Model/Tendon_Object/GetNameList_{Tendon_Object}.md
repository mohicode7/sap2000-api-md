# GetNameList

## Syntax

SapObject.SapModel.TendonObj.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of tendon object names retrieved by the program.

MyName

This is a one-dimensional array of tendon object names. The MyName array is created as a dynamic, zero-based, array by the APIuser:



```
   Dim MyName() as String

 
```



The array is dimensioned to (NumberNames – 1) inside the Sap2000 program, filled with the names, and returned to the APIuser.

## Remarks

This function retrieves the names of all defined tendon objects.

The function returns zero if the names are successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetTendonObjectNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberNames As Long

      Dim MyName() As String

      Dim NumberPoints As Long

      Dim MyType() As Long

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name1)

      ret = SapModel.TendonObj.AddByPoint("2", "8", Name2)



   'set tendon geometry

      NumberPoints = 3

      ReDim MyType(NumberPoints - 1)

      ReDim x(NumberPoints - 1)

      ReDim y(NumberPoints - 1)

      ReDim z(NumberPoints - 1)

      MyType(0) = 1

      MyType(1) = 7

      MyType(2) = 6

      x(0) = 0:  y(0) = 0

      x(1) = 288:  y(1) = -12

      x(2) = 576:  y(2) = 0

      ret = SapModel.TendonObj.SetTendonData(Name1, NumberPoints, MyType, x, y, z, "Local")

      ret = SapModel.TendonObj.SetTendonData(Name2, NumberPoints, MyType, x, y, z, "Local")



   'get tendon object names

      ret = SapModel.TendonObj.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
