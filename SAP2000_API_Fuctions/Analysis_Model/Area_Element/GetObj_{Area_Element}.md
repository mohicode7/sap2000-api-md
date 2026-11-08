# GetObj

## Syntax

Sap2000.AreaElm.GetObj

## VB6 Procedure

Function GetObj(ByVal Name As String, ByRef Obj As String) As Long

## Parameters

Name

The name of an existing area element.

Obj

The name of the area object from which the area element was created.

## Remarks

This function retrieves the name of the area object from which an area element was created.

The function returns zero if the information is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetObjForAreaElm()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign auto mesh options

      ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 3, 3, , , , , , , , , , , , , , , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get object information for an area element

      ret = SapModel.AreaElm.GetObj("3-2", Obj)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
