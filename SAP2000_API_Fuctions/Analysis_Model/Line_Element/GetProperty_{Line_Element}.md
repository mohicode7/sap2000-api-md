# GetProperty

## Syntax

SapObject.SapModel.LineElm.GetProperty

## VB6 Procedure

Function GetProperty(ByVal Name As String, ByRef PropName As String, ByRef ObjType As Long, ByRef Var As Boolean, ByRef sVarRelStartLoc As Double, sVarTotalLength As Double) As Long

## Parameters

Name

The name of an existing line element.

PropName

The name of the frame section, cable or tendon property assigned to the line element.

ObjType

This is 0, 1, 2 or 3, indicating the type of object from which the line element was created.

0 = Straight frame object

1 = Curved frame object

2 = Cable object

3 = Tendon object

Var

This item is True if the specified property is a nonprismatic (variable) frame section property.

sVarTotalLength

This is the total assumed length of the nonprismatic section. A zero value for this item means that the section length is the same as the line element length.

sVarRelStartLoc

This is the relative distance along the nonprismatic section to the I-End (start) of the line element. This item is ignored when the sVarTotalLengthitem is 0.

## Remarks

This function retrieves the property assignment to a line element.

The function returns zero if the property data is successfully retrieved, otherwise it returns a nonzero value.

The sVarTotalLength and sVarRelStartLoc items apply only when the Var item is True.

## VBA Example



```
Sub GetLineElementProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String

      Dim ObjType As Long

      Dim Var As Boolean

      Dim sVarRelStartLoc As Double

      Dim sVarTotalLength As Double



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



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get property information for a line element

      ret = SapModel.LineElm.GetProperty("3-1", PropName, ObjType, Var, sVarRelStartLoc, sVarTotalLength)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
