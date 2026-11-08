# SetTrapezoidal

## Syntax

SapObject.SapModel.PropFrame.SetTrapezoidal

## VB6 Procedure

Function SetTrapezoidal(ByVal Name As String, ByVal MatProp As String, ByVal t3 As Double, ByVal t2 As Double, ByVal t2b As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

t2

The section top width. [L]

t2b

The section bottom width. [L]

Color

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function initializes a solid trapezoidal frame section property. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropTrapezoidal()
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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



  'set new frame section property

      ret = SapModel.PropFrame.SetTrapezoidal("R1", "4000Psi", 20, 20, 12)



  'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 17.2.0

## See Also

GetTrapezoidal

SetRebarBeam

SetRebarColumn

GetRebarBeam

GetRebarColumn
