# GetLoadStrain

## Syntax

SapObject.SapModel.LineElm.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems As Long, ByRef LineName() As String, ByRef LoadPat() As String, ByRef DOF() As Long, ByRef Val() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing line object, line element or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

The total number of strain loads retrieved for the specified line elements.

LineName

This is an array that includes the name of the line element associated with each strain load.

LoadPat

This is an array that includes the name of the load pattern associated with each strain load.

DOF

This is an array that includes 1, 2, 3, 4, 5 or 6, indicating the degree of freedom associated with each strain load.

1 = Strain11

2 = Strain12

3 = Strain13

4 = Curvature1

5 = Curvature2

6 = Curvature3

Val

This is an array that includes the strain value. [L/L] for DOF = 1, 2 and 3 and [1/L] for DOF = 4, 5 and 6

PatternName

This is an array that includes the joint pattern name, if any, used to specify the strain load.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are retrieved for the line elements corresponding to the line object specified by the Name item.

If this item is Element, the load assignments are retrieved for the line element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the line elements corresponding to all line objects included in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for line elements corresponding to all selected line objects, and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments to line elements.

The function returns zero if the strain load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim LineName() As String

      Dim LoadPat() As String

      Dim DOF() As Long

      Dim Val() As Double

      Dim PatternName() As String



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



   'assign frame object strain load

      ret = SapModel.FrameObj.SetLoadStrain("1", "DEAD", 1, 0.001)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element strain load

      ret = SapModel.LineElm.GetLoadStrain("1-1", NumberItems, LineName, LoadPat, DOF, Val, PatternName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also
