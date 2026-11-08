# GetObj

## Syntax

Sap2000.LineElm.GetObj

## VB6 Procedure

Function GetObj(ByVal Name As String, ByRef Obj As String, ByRef ObjType As Long, ByRef RDI As Double, RDJ As Double) As Long

## Parameters

Name

The name of an existing line element.

Obj

The name of the frame, cable or tendon object from which the line element was created.

ObjType

This is 0, 1, 2 or 3, indicating the type of object from which the line element was created.

0 = Straight frame object

1 = Curved frame object

2 = Cable object

3 = Tendon object

RDI

The relative distance from the I-End of the object identified by the Obj item to the I-End of the considered line element. The relative distance is calculated as the distance from the I-End of the object to the I-End of the line element divided by the length of the object.

RDJ

The relative distance from the I-End of the object identified by the Obj item to the J-End of the considered line element. The relative distance is calculated as the distance from the I-End of the object to the J-End of the line element divided by the length of the object.

## Remarks

This function retrieves information about the object from which a line element was created.

The function returns zero if the information is successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetObjForLineElm()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Obj As String

      Dim ObjType As Long

      Dim RDI As Double

      Dim RDJ As Double



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



   'assign auto mesh options

      ret = SapModel.FrameObj.SetAutoMesh("ALL", True, True, True, 2, 0, Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get object information for a line element

      ret = SapModel.LineElm.GetObj("3-1", Obj, ObjType, RDI, RDJ)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
