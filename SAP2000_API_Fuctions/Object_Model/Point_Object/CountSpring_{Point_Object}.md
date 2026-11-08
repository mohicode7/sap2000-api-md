# CountSpring

## Syntax

SapObject.SapModel.PointObj.CountSpring

## VB6 Procedure

Function CountSpring() As Long

## Parameters

None

## Remarks

This function returns the total number of point objects in the model with spring assignments.

## VBA Example



```
Sub CountSpringSupportedPointObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim i As Long

      Dim k() As Double



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

      redim k(5)

      For i = 0 to 5

         k(i) = i + 1

      Next i

      ret = SapModel.PointObj.SetSpring("3", k)



   'get number of point objects with spring assignments

      Count = SapModel.PointObj.CountSpring



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BPoint_Object%7D.md)

[SetSpring](SetSpring.md)

[DeleteSpring](DeleteSpring_%7BPoint_Object%7D.md)
