# CountSpring

## Syntax

SapObject.SapModel.PointElm.CountSpring

## VB6 Procedure

Function CountSpring() As Long

## Parameters

None

## Remarks

This function returns the total number of point elements in the model with spring assignments.

## VBA Example



```
Sub CountPointElementsWithSprings()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim i as Long

      Dim k() as Double

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



   'add joint spring assignment

      ReDim k(5)

      For i = 0 to 5

         k(i) = i + 1

      Next i

      ret = SapModel.PointObj.SetSpring("3", k)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get number of point elements with springs

      Count = SapModel.PointElm.CountSpring



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BPoint_Element%7D.md)

[GetSpringCoupled](GetSpringCoupled_%7BPoint_Element%7D.md)
