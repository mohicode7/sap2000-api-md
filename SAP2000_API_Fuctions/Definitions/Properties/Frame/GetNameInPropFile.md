# GetNameInPropFile

## Syntax

SapObject.SapModel.PropFrame.GetNameInPropFile

## VB6 Procedure

Function GetNameInPropFile(ByVal Name As String, ByRef NameInFile As String, ByRef FileName As String, ByRef MatProp As String, ByRef PropType As eFramePropType) As Long

## Parameters

Name

The name of an existing frame section property.

NameInFile

The name of the specified frame section property in the frame section property file.

FileName

The name of the frame section property file from which the specified frame section property was obtained.

MatProp

The name of the material property for the section.

PropType

This is one of the following items in the eFramePropType enumeration.

SECTION\_I = 1

SECTION\_CHANNEL = 2

SECTION\_T = 3

SECTION\_ANGLE = 4

SECTION\_DBLANGLE = 5

SECTION\_BOX = 6

SECTION\_PIPE = 7

SECTION\_RECTANGULAR = 8

SECTION\_CIRCLE = 9

SECTION\_GENERAL = 10

SECTION\_DBCHANNEL = 11

SECTION\_AUTO = 12

SECTION\_SD = 13

SECTION\_VARIABLE = 14

SECTION\_JOIST = 15

SECTION\_BRIDGE = 16

SECTION\_COLD\_C = 17

SECTION\_COLD\_2C = 18

SECTION\_COLD\_Z = 19

SECTION\_COLD\_L = 20

SECTION\_COLD\_2L = 21

SECTION\_COLD\_HAT = 22

SECTION\_BUILTUP\_I\_COVERPLATE = 23

SECTION\_PCC\_GIRDER\_I = 24

SECTION\_PCC\_GIRDER\_U = 25

## Remarks

This function retrieves the names of the section property file from which an imported frame section originated, and it also retrieves the section name used in the property file.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

If the specified frame section property was not imported, blank strings are returned for NameInFile and FileName.

## VBA Example



```
Sub GetFramePropNameInFile()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NameInFile As String

          Dim FileName As String

          Dim MatProp As String

          Dim PropType As eFramePropType



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

          ret = SapModel.PropFrame.ImportProp("MyFrame", "A992Fy50", "AISC16.XML", "W18X35")



       'get frame property name in file

          ret = SapModel.PropFrame.GetNameInPropFile("MyFrame", NameInFile, FileName, MatProp, PropType)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.03.

## See Also
