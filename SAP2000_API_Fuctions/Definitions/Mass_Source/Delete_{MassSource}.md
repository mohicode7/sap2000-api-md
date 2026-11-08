# Delete {Mass Source}

## Syntax

SapObject.SapModel.SourceMass.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of the mass source to be deleted.

## Remarks

This function deletes an existing mass source.

The function returns zero if the mass source is successfully deleted; otherwise it returns a nonzero value. If the mass source to be deleted is the default mass source, a nonzero value is returned and th mass source is not deleted.

## VBA Example



```
Sub DeleteMassSource()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

     Dim LoadPat(0) As String
      Dim SF(0) As Double
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
   'add a new mass source and make it the default mass source
      LoadPat(0) = "DEAD"
      SF(0) = 1.25
      ret = SapModel.SourceMass.SetMassSource("MyMassSource", True, True, True, True, 1, LoadPat, SF)
   'delete mass source MSSSRC1

      ret = SapModel.SourceMass.Delete("MSSSRC1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also
