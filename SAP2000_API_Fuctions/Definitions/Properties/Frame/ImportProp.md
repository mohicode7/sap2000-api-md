# ImportProp

## Syntax

SapObject.SapModel.PropFrame.ImportProp

## VB6 Procedure

Function ImportProp(ByVal Name As String, ByVal MatProp As String, ByVal FileName As String, ByVal PropName As String, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added. This name does not need to be the same as the PropName item.

MatProp

The name of the material property for the section.

FileName

The name of the frame section property file from which to get the frame section property specified by the PropName item.

In most cases you can input just the name of the property file (e.g. AISC16.XML) and the program will be able to find it. In some cases you may have to input the full path to the property file.

PropName

The name of the frame section property, inside the property file specified by the FileName item, that is to be imported.

Color

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function imports a frame section property from a property file.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

If the property file is not found, or the specified property name is not found in the property file, the section is set to be a general section with default properties.

## VBA Example



```
Sub ImportFrameProp()

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



       'import new frame section property

          ret = SapModel.PropFrame.ImportProp("W18X35", "A992Fy50", "AISC16.XML", "W18X35")



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetPropFileNameList](GetPropFileNameList.md)
