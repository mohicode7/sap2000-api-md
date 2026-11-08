# GetActiveDOF

## Syntax

SapObject.SapModel.Analyze.GetActiveDOF

## VB6 Procedure

Function GetActiveDOF(ByRef DOF() As Boolean) As Long

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

This function retrieves the model global degrees of freedom.

The function returns zero if the degrees of freedom are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetModelActiveDOF()

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



   'get model degrees of freedom

      ret = SapModel.Analyze.GetActiveDOF(DOF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetActiveDOF](SetActiveDOF.md)
