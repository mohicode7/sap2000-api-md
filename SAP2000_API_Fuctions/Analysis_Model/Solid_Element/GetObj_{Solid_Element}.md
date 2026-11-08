# GetObj

## Syntax

Sap2000.SolidElm.GetObj

## VB6 Procedure

Function GetObj(ByVal Name As String, ByRef Obj As String) As Long

## Parameters

Name

The name of an existing solid element.

Obj

The name of the solid object from which the solid element was created.

## Remarks

This function retrieves the name of the solid object from which a solid element was created.

The function returns zero if the information is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetObjForSolidElm()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign auto mesh options

      ret = SapModel.SolidObj.SetAutoMesh("ALL", 1, 2, 2, 2, , , , , , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get object information for an solid element

      ret = SapModel.SolidElm.GetObj("3-2", Obj)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
