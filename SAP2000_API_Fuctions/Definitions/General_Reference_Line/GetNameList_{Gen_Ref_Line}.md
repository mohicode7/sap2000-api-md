# GetNameList

## Syntax

SapObject.SapModel.GenRefLine.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of general reference line names retrieved by the program.

MyName

This is a one-dimensional array of general reference line names. The MyName array is created as a dynamic, zero-based array by the API user:



```
Dim MyName() As String

 
```



The array is dimensioned to (NumberNames – 1) inside the SAP2000 program, filled with the names, and returned to the API user.

## Remarks

This function retrieves the names of all defined general reference lines.

The function returns the names; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetGenRefLineNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'define new general reference line

      ret = SapModel.GenRefLine.SetLine("GRef1", 120, 5)



   'get general reference line names

      ret = SapModel.GenRefLine.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.0.

## See Also

[SetLine](SetLine_%7BGen_Ref_Line%7D.md)
