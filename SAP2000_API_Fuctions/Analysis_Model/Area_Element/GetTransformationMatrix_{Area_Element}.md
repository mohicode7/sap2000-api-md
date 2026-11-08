# GetTransformationMatrix

## Syntax

Sap2000.AreaElm.GetTransformationMatrix

## VB6 Procedure

Function GetTransformationMatrix(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing area element.

Value

Value is an array of nine direction cosines that define the transformation matrix.

The following matrix equation shows how the transformation matrix is used to convert items from the area element local coordinate system to the global coordinate system.

![](../../../source-assets/assets/images/Images/Transformation%20Matrix.JPG)

In the equation c0 through c8 are the nine values from the transformation array, (Local1, Local2, Local3) are an item (such as a load) in the element local coordinate system, and (GlobalX, GlobalY, GlobalZ) are the same item in the global coordinate system.

## Remarks

The function returns zero if the transformation matrix is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaElementMatrix()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value() As Double



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



   'assign area object local axis angle

      ret = SapModel.AreaObj.SetLocalAxes("3", 30)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get area element transformation matrix

      redim Value(8)

      ret = SapModel.AreaElm.GetTransformationMatrix("3", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
