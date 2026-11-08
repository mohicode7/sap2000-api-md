# Add

## Syntax

SapObject.SapModel.GDispl.Add

## VB6 Procedure

Function Add(ByVal Name As String, ByVal MyType As Long) As Long

## Parameters

Name

The name of a new generalized displacement.

MyType

This is 1 or 2 indicating the generalized displacement type.

1 = Translational

2 = Rotational

## Remarks

This function adds a new generalized displacement with the specified name and type.

The function returns zero if the generalized displacement is successfully added, otherwise it returns a nonzero value.

The new generalized displacement must have a different name from all other generalized displacements. If the name is not unique, an error will be returned.

## VBA Example



```
Sub AddNewGDispl()

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

      ret = SapModel.GDispl.Add("GD1", 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPoint](SetPoint.md)
