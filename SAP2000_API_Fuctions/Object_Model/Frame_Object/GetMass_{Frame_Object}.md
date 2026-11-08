# GetMass

## Syntax

SapObject.SapModel.FrameObj.GetMass

## VB6 Procedure

Function GetMass(ByVal Name As String, ByRef MassOverL As Double) As Long

## Parameters

Name

The name of an existing frame object.

MassOverL

The mass per unit length assigned to the frame object. [M/L]

## Remarks

This function retrieves the frame mass per unit length assignment for frame objects.

The function returns zero if the mass assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MassOverL As Double



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



   'assign frame mass

      ret = SapModel.FrameObj.SetMass("ALL", .0001, False, Group)



   'get frame mass assignment

      ret = SapModel.FrameObj.GetMass("3", MassOverL)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMass](SetMass_%7BFrame_Object%7D.md)

[DeleteMass](DeleteMass_%7BFrame_Object%7D.md)
