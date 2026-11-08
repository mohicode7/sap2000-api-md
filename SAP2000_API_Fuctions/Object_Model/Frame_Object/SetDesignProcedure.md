# SetDesignProcedure

## Syntax

SapObject.SapModel.FrameObj.SetDesignProcedure

## VB6 Procedure

Function SetDesignProcedure(ByVal Name As String, ByVal MyType As Long, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

MyType

This is 1 or 2, indicating the design procedure type desired for the specified frame object.

1 = Default from material

2 = No design

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

Selection = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is Selection, assignment is made to all selected frame objects and the Name item is ignored.

## Remarks

This function sets the design procedure for frame objects.

The function returns zero if the design procedure is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameObjDesignProcedure()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'set design procedure

      ret = SapModel.FrameObj.SetDesignProcedure("8", 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetDesignProcedure](GetDesignProcedure.md)
