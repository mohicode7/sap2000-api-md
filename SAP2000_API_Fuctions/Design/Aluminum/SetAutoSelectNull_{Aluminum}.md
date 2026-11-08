# SetAutoSelectNull

## Syntax

SapObject.SapModel.DesignAluminum.SetAutoSelectNull

## VB6 Procedure

Function SetAutoSelectNull(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function removes the auto select section assignments from all specified frame objects that have a aluminum frame design procedure.

The function returns zero if the auto select section assignments are successfully removed; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAluminumAutoSelectSectionsNull()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MyName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add aluminum material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)



   'create new aluminum frame section properties

      ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)

      ret = SapModel.PropFrame.SetISection("AI2", Name , 18, 6, 0.6, 0.3, 6, 0.6)

      ret = SapModel.PropFrame.SetISection("AI3", Name , 18, 6, 0.7, 0.3, 6, 0.7)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



   'define new auto select list frame section property

      ReDim MyName(2)

      MyName(0) = "AI"

      MyName(1) = "AI2"

      MyName(2) = "AI3"

      ret = SapModel.PropFrame.SetAutoSelectAluminum("AUTO1", 3, MyName)



   'set frame section properties

      ret = SapModel.FrameObj.SetSection("8", "AUTO1")

      ret = SapModel.FrameObj.SetSection("10", "AUTO1")



   'set frame object selected

      ret = SapModel.FrameObj.SetSelected("8", True)



   'set auto select section null

      ret = SapModel.DesignAluminum.SetAutoSelectNull("",SelectedObjects)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
