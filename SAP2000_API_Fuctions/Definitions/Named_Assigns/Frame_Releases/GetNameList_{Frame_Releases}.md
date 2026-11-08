# GetNameList

## Syntax

SapObject.SapModel.NamedAssign.ReleaseFrame.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of frame end release names retrieved by the program.

MyName

This is a one-dimensional array of frame end release names. The MyName array is created as a dynamic, zero-based array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

## Remarks

This function retrieves the names of all defined frame end releases.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero values.

## VBA Example



```
Sub GetFrameEndReleaseNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ii() As Boolean

      Dim jj() As Boolean

      Dim StartValue() As Double

      Dim EndValue() As Double

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



   'add new end release

      ReDim ii(5)

      ReDim jj(5)

      ReDim StartValue(5)

      ReDim EndValue(5)

      ii(0) = True

      StartValue(0) = 10

      ret = SapModel.NamedAssign.ReleaseFrame.SetReleases("FREL1", ii, jj, StartValue, EndValue)

      ii(0) = False

      StartValue(0) = 0

      ii(5) = True

      jj(5) = True

      ret = SapModel.NamedAssign.ReleaseFrame.SetReleases("FREL2", ii, jj, StartValue, EndValue)



   'get frame end release names

      ret = SapModel.NamedAssign.ReleaseFrame.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
