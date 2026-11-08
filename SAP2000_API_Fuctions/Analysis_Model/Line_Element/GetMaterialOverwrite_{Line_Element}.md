# GetMaterialOverwrite

## Syntax

SapObject.SapModel.LineElm.GetMaterialOverwrite

## VB6 Procedure

Function GetMaterialOverwrite(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined line element.

PropName

This is None, indicating that no material overwrite exists for the specified line element, or it is the name of an existing material property.

## Remarks

This function retrieves the material overwrite assigned to a line element, if any. It returns None if there is no material overwrite assignment.

The function returns zero if the material overwrite assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmMaterialOverwrite()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String



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



   'assign material overwrite to frame object

      ret = SapModel.FrameObj.SetMaterialOverwrite("1", "4000Psi")



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get material overwrite assignment for line element

      ret = SapModel.LineElm.GetMaterialOverwrite("1-1", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
