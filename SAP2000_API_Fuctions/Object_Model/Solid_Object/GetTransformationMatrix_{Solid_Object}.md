# GetTransformationMatrix

## Syntax

Sap2000.SolidObj.GetTransformationMatrix

## VB6 Procedure

Function GetTransformationMatrix(ByVal Name As String, ByRef Value() As Double,Optional ByVal IsGlobal As Boolean = True) As Long

## Parameters

Name

The name of an existing solid object.

Value

Value is an array of nine direction cosines that define the transformation matrix.

The following matrix equation shows how the transformation matrix is used to convert items from the solid object local coordinate system to the global coordinate system.

![](../../../source-assets/assets/images/Images/Transformation%20Matrix.JPG)

In the equation, c0 through c8 are the nine values from the transformation array, (Local1, Local2, Local3) are an item (such as a load) in the object local coordinate system, and (GlobalX, GlobalY, GlobalZ) are the same item in the global coordinate system.

The transformation from the local coordinate system to the present coordinate system is the same as that shown above for the global system if you substitute the present system for the global system.

IsGlobal

If this item is True, the transformation matrix is between the Global coordinate system and the solid object local coordinate system.

If this item is False, the transformation matrix is between the present coordinate system and the solid object local coordinate system.

## Remarks

The function returns zero if the solid object transformation matrix is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidObjectMatrix()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'get solid object transformation matrix

      ReDim Value(8)

      ret = SapModel.SolidObj.GetTransformationMatrix("3", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
