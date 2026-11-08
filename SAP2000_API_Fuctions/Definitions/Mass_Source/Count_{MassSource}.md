# Count {Mass Source}

## Syntax

SapObject.SapModel.SourceMass.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the number of defined mass sources.

## VBA Example



```
Sub CountMassSources()

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
 
   'add a new mass source and make it the default mass sourc
      LoadPat(0) = "DEAD"
      SF(0) = 1.25
      ret = SapModel.SourceMass.SetMassSource("MyMassSource", True,       True, True, True, 1, LoadPat, SF)



   'get mass source count

      Count = SapModel.SourceMass.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also
