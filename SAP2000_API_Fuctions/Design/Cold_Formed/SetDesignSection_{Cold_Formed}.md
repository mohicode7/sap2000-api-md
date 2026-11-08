# SetDesignSection

## Syntax

SapObject.SapModel.DesignColdFormed.SetDesignSection

## VB6 Procedure

Function SetDesignSection(ByVal Name As String, ByVal PropName As String, ByVal LastAnalysis As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

PropName

The name of an existing frame section property to be used as the design section for the specified frame objects. This item applies only when LastAnalysis = False.

LastAnalysis

If this item is True, the design section for the specified frame objects is reset to the last analysis section for the frame object. If it is False, the design section is set to that specified by PropName.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function modifies the design section for all specified frame objects that have a cold formed frame design procedure.

The function returns zero if the design section is successfully modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetColdFormedDesignSection()

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)

      ret = SapModel.PropFrame.SetColdC("CdC2", Name , 9, 3, 0.07, 0.25, 0.5)

      ret = SapModel.PropFrame.SetColdC("CdC3", Name , 9, 3, 0.08, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'define new auto select list frame section property

      ReDim MyName(2)

      MyName(0) = "CdC"

      MyName(1) = "CdC2"

      MyName(2) = "CdC3"

      ret = SapModel.PropFrame.SetAutoSelectColdFormed("AUTO1", 3, MyName)



   'set frame section properties

      ret = SapModel.FrameObj.SetSection("8", "AUTO1")

      ret = SapModel.FrameObj.SetSection("10", "AUTO1")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start cold formed design

      ret = SapModel.DesignColdFormed.StartDesign



   'set design section

      ret = SapModel.DesignColdFormed.SetDesignSection("8", "CdC3", False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetDesignSection](GetDesignSection_%7BCold_Formed%7D.md)
