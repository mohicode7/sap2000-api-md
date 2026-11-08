# ColumnDesignForces

## Syntax

SapObject.SapModel.DesignResults.DesignForces.ColumnDesignForces

## VB6 Procedure

Function ColumnDesignForces (

ByVal Name As String,

ByRef NumberResults As Long,

ByRef FrameName As String(),

ByRef ComboName As String (),

ByRef Station As Double (),

ByRef P As Double(),

ByRef V2 As Double(),

ByRef V3 As Double(),

ByRef T As Double(),

ByRef M2 As Double(),

ByRef M3 As Double(),

Optional ItemType As eItemType = eItemType.Objects) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the *ItemType* item.

NumberResults

The number of results, this is the length of all output arrays

FrameName

The frame unique name associated with each result

ComboName

The load combination for which the results are reported

Station

The location, measured from the I-end of the frame object, where the results are reported [L]

P

The design axial force in the frame object local 1-axis direction [F]

V2

The design shear force in the frame object local 2-axis direction [F]

V3

The design shear force in the frame object local 3-axis direction [F]

T

The design torsional moment about the frame object local 1-axis [FL]

M2

The design bending moment about the frame object local 2-axis [FL]

M3

The design bending moment about the frame object local 3-axis [FL]

ItemType [Optional]

This is one of the following items in the eItemType enumeration:

- Object = 0
- Group = 1
- SelectedObjects = 2

If this item is Objects, results are retrieved for the frame object specified by the Name item. This is the default.

If this item is Group, results are retrieved for all frame objects specified by the Name item.

If this item is SelectedObjects, results are retrieved for all selected frame objects and the Name item is ignored.

## Remarks

This function retrieves design forces for an existing designed column.

The function returns zero if the forces are successfully recovered, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColumnDesignForces()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim NumberResults As Long
      Dim FrameName() As String
      Dim ComboName() As String
      Dim Station() As Double
      Dim P() As Double
      Dim V2() As Double
      Dim V3() As Double
      Dim T() As Double
      Dim M2() As Double
      Dim M3() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(e2DFrameType_PortalFrame, 3, 124, 3, 200)



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'run design

      ret = SapModel.DesignSteel.StartDesign



   'get column design forces

      ret = SapModel.DesignResults.DesignForces.ColumnDesignForces(“1”, NumberResults, FrameName, ComboName, Station, P, V2, V3, T, M2, M3, eItemType_Objects)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.2
