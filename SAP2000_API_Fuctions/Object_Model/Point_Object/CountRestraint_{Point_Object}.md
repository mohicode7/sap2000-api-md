# CountRestraint

## Syntax

SapObject.SapModel.PointObj.CountRestraint

## VB6 Procedure

Function CountRestraint() As Long

## Parameters

None

## Remarks

This function returns the total number of point objects in the model with restraint assignments.

## VBA Example



```
Sub CountRestrainedPointObjects()

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



   'get number of restrained point objects

      Count = SapModel.PointObj.CountRestraint



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetRestraint](GetRestraint_%7BPoint_Object%7D.md)

[SetRestraint](SetRestraint.md)

[DeleteRestraint](DeleteRestraint.md)
