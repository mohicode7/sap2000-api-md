# SetSDSection

## Syntax

SapObject.SapModel.PropFrame.SetSDSection

## VB6 Procedure

Function SetSDSection(ByVal Name As String, ByVal MatProp As String, Optional ByVal DesignType As Long = 0, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

MatProp

The name of the base material property for the section.

DesignType

This is 0, 1, 2 or 3, indicating the design option for the section.

0 = No design

1 = Design as general steel section

2 = Design as a concrete column; check the reinforcing

3 = Design as a concrete column; design the reinforcing

When DesignType = 1 is assigned, the material property specified by the MatProp item must be a steel material; otherwise the program sets DesignType = 0.

Similarly, when DesignType = 2 or DesignType = 3 is assigned, the material property specified by the MatProp item must be a concrete material; otherwise the program sets DesignType = 0.

Color

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function initializes a section designer property.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropSDSection()

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

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000Psi")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetSDSection](GetSDSection.md)
