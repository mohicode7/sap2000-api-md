# SetPattern

## Syntax

SapObject.SapModel.PatternDef.SetPattern

## VB6 Procedure

Function SetPattern(ByVal Name As String) As Long

## Parameters

Name

The name of a new joint pattern.

## Remarks

The function defines a new joint pattern. It returns an error if the specified joint pattern name is already in use.

The function returns zero if the new joint pattern is successfully defined, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddPattern()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'add joint pattern

      ret = SapModel.PatternDef.SetPattern("MyPattern")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
