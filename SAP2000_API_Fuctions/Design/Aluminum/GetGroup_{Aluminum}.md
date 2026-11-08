# GetGroup

## Syntax

SapObject.SapModel.DesignAluminum.GetGroup

## VB6 Procedure

Function GetGroup(ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of groups selected for aluminum design.

MyName

This is an array that includes the name of each group selected for aluminum design.

## Remarks

This function retrieves the names of all groups selected for aluminum design.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignGroup()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

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



   'create new aluminum frame section property

      ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



   'select group for aluminum design

      ret = SapModel.DesignAluminum.SetGroup("ALL", True)



   'get groups selected for aluminum design

      ret = SapModel.DesignAluminum.GetGroup(NumberItems, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetGroup](SetGroup_%7BAluminum%7D.md)
