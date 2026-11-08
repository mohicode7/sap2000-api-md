# SetTypeOAPI

## Syntax

SapObject.SapModel.GDispl.SetTypeOAPI

## VB6 Procedure

Function SetTypeOAPI(ByVal Name As String, ByVal MyType As Long) As Long

## Parameters

Name

The name of an existing generalized displacement.

MyType

This is 1 or 2, indicating the generalized displacement type.

1 = Translational

2 = Rotational

## Remarks

This function sets the generalized displacement type.

The function returns zero if the type is successfully set, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetGDisplType()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add generalized displacement

         ret = SapModel.GDispl.Add("GD1", 2)



   'modify generalized displacement type

      ret = SapModel.GDispl.SetTypeOAPI("GD1", 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed function name to SetTypeOAPI in v17.0.0.

## See Also

[GetTypeOAPI](GetType_%7BGeneralized_Displacement%7D.md)
