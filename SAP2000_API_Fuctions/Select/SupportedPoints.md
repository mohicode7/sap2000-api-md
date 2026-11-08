# SupportedPoints

## Syntax

SapObject.SapModel.SelectObj.SupportedPoints

## VB6 Procedure

Function SupportedPoints(ByRef DOF() As Boolean, Optional ByVal CSys As String = "Local", Optional ByVal DeSelect As Boolean = False, Optional ByVal SelectRestraints As Boolean = True, Optional ByVal SelectJointSprings As Boolean = True, Optional ByVal SelectLineSprings As Boolean = True, Optional ByVal SelectAreaSprings As Boolean = True, Optional ByVal SelectSolidSprings As Boolean = True, Optional ByVal SelectOneJointLinks As Boolean = True) As Long

## Parameters

DOF

This is an array of six booleans for the six degrees of freedom of a point object.

DOF(0) = U1

DOF(1) = U2

DOF(2) = U3

DOF(3) = R1

DOF(4) = R2

DOF(5) = R3

CSys

The name of the coordinate system in which degrees of freedom (DOF) are specified. This is either Local or the name of a defined coordinate system. Local means the point local coordinate system.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

SelectRestraints

If this item is True then points with restraint assignments in one of the specified degrees of freedom are selected or deselected.

SelectJointSprings

If this item is True then points with joint spring assignments in one of the specified degrees of freedom are selected or deselected.

SelectLineSprings

If this item is True, points with a contribution from line spring assignments in one of the specified degrees of freedom are selected or deselected.

SelectAreaSprings

If this item is True, points with a contribution from area spring assignments in one of the specified degrees of freedom are selected or deselected.

SelectSolidSprings

If this item is True, points with a contribution from solid surface spring assignments in one of the specified degrees of freedom are selected or deselected.

SelectOneJointLinks

If this item is True, points with one joint link assignments in one of the specified degrees of freedom are selected or deselected.

## Remarks

This function selects or deselects point objects with support in the specified degrees of freedom.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectSupportedPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'select supported points

      Redim DOF(5)

      DOF(2) = True

      ret = SapModel.SelectObj.SupportedPoints(DOF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
