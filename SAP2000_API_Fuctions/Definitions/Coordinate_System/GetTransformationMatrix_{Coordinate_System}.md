# GetTransformationMatrix

## Syntax

SapObject.SapModel.CoordSys.GetTransformationMatrix

## VB6 Procedure

Function GetTransformationMatrix(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing coordinate system.

Value

Value is an array of nine direction cosines that define the transformation matrix from the specified global coordinate system to the global coordinate system.

The following matrix equation shows how the transformation matrix is used to convert coordinates from a coordinate system to the global coordinate system.

![](../../../source-assets/assets/images/Images/Transformation%20Matrix.JPG)

In the equation, c0 through c8 are the nine values from the transformation array; (x, y, z) are the coordinates of a point in the CSys coordinate system; (ux, uy, uz) are the offset of the origin of the CSys coordinate system from the global coordinate system; and (gx, gy, gz) are the global coordinates of the point.

## Remarks

The function returns zero if the coordinate system transformation matrix is successfully returned, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCoordSystemMatrix()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim Value() as double



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



   'define new coordinate system

      ret = SapModel.CoordSys.SetCoordSys("CSys1", 1000, 1000, 0, 0, 0, 0)



   'get coordinate system transformation matrix

      redim Value(8)

      ret = SapModel.CoordSys.GetTransformationMatrix("CSys1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
