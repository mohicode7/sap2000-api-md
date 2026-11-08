# GetTransformationMatrix

## Syntax

Sap2000.FrameObj.GetTransformationMatrix

## VB6 Procedure

Function GetTransformationMatrix(ByVal Name As String, ByRef Value() As Double,Optional ByVal IsGlobal As Boolean = True) As Long

## Parameters

Name

The name of an existing frame object.

Value

Value is an array of nine direction cosines that define the transformation matrix.

The following matrix equation shows how the transformation matrix is used to convert items from the frame object local coordinate system to the global coordinate system.

![](../../../source-assets/assets/images/Images/Transformation%20Matrix.JPG)

In the equation, c0 through c8 are the nine values from the transformation array, (Local1, Local2, Local3) are an item (such as a load) in the object local coordinate system, and (GlobalX, GlobalY, GlobalZ) are the same item in the global coordinate system.

The transformation from the local coordinate system to the present coordinate system is the same as that shown above for the global system if you substitute the present system for the global system.

IsGlobal

If this item is True, the transformation matrix is between the Global coordinate system and the frame object local coordinate system.

If this item is False, the transformation maxtrix is between the present coordinate system and the frame object local coordinate system.

## Remarks

The function returns zero if the frame object transformation matrix is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameObjectMatrix()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'assign frame object local axis angle

      ret = SapModel.FrameObj.SetLocalAxes("3", 30)



   'get frame object transformation matrix

      ReDim Value(8)

      ret = SapModel.FrameObj.GetTransformationMatrix("3", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
