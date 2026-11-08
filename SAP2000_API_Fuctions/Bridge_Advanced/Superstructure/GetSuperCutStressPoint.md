# GetSuperCutStressPoint

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.GetSuperCutStressPoint

## VB6 Procedure

Function GetSuperCutStressPoint(ByVal Name As String, ByVal CutIndex As Long, ByVal PointIndex As Long, ByRef X As Double, ByRef Y As Double, ByRef MatProp As String, ByRef Note as String) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function CountSuperCut.

PointIndex

The index number of the stress point in this section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function CountSuperCutStressPoint.

X, Y

The transverse and vertical coordinates of the stress point in the section, measured from the bottom left corner of the section.  X is positive to the right when looking upstation. Y is positive upward. [L]

MatProp

The name of the material property at this stress point.

Note

A description of the stress point that may be used for identification. Points that are pre-defined by the program will have prescribed notes.

## Remarks

This function returns location and material information about a single stress point at a superstructure section cut in a bridge object.

The function returns zero if the information is successfully retrieved, otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutStressPointInfo()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim CountPoint As Long

      Dim X As Double,

      Dim Y As Double

      Dim MatProp As String

      Dim Note As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\MyBridge.sdb"

      ret = SapModel.File.OpenFile(FileName)



   'get section cut count

      ret = SapModel.BridgeAdvancedSuper.CountSuperCut("BOBJ1", Count)



   'get section cut stress point count

      ret = SapModel.BridgeAdvancedSuper.CountSuperCutStressPoint("BOBJ1", 1, CountPoint)



   'get section cut stress point location information

      ret = SapModel.BridgeAdvancedSuper.GetSuperCutStressPoint("BOBJ1", 1, 1, X, Y, MatProp, Note)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[CountSuperCut](CountSuperCut.md)

[CountSuperCutStressPoint](GetSuperCutStressPoint.md)
