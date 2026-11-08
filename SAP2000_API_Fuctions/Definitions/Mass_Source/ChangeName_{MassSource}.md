# Change Name {Mass Source}

## Syntax

SapObject.SapModel.SourceMass.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The name of an existing mass source.

NewName

The new name for the mass source.

## Remarks

This function changes the name of an existing mass source.

The function returns zero if the mass source is successfully changed, otherwise it returns a nonzero value. If the new name already exists, a nonzero value is returned and the mass source name is not changed.

## VBA Example



```
Sub ChangeMassSourceName()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'change mass source name from MSSSRC1 to MyMassSource

      ret = SapModel.SourceMass.ChangeName("MSSSRC1", "MyMassSource")

   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0..

## See Also
