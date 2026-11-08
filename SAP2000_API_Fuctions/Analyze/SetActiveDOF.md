# SetActiveDOF

## Syntax

SapObject.SapModel.Analyze.SetActiveDOF

## VB6 Procedure

Function SetActiveDOF(ByRef DOF() As Boolean) As Long

## Parameters

DOF

This is an array of 6 boolean values, indicating if the specified model global degree of freedom is active.

DOF(0) = UX

DOF(1) = UY

DOF(2) = UZ

DOF(3) = RX

DOF(4) = RY

DOF(5) = RZ

## Remarks

This function sets the model global degrees of freedom.

The function returns zero if the degrees of freedom are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetModelActiveDOF()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean



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



   'set model degrees of freedom

      ReDim DOF(5)

      DOF(0) = True

      DOF(1) = True

      DOF(2) = True

      ret = SapModel.Analyze.SetActiveDOF(DOF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetActiveDOF](GetActiveDOF.md)
