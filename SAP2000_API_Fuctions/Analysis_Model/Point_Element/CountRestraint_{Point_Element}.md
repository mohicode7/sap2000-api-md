# CountRestraint

## Syntax

SapObject.SapModel.PointElm.CountRestraint

## VB6 Procedure

Function CountRestraint() As Long

## Parameters

None

## Remarks

This function returns the total number of point elements in the model with restraint assignments.

## VBA Example



```
Sub CountRestrainedPointElements()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim Count as Long



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



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get number of restrained point elements

      Count = SapModel.PointElm.CountRestraint



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetRestraint](GetRestraint_%7BPoint_Element%7D.md)
