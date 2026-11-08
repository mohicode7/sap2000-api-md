# GetGroup

## Syntax

SapObject.SapModel.DesignColdFormed.GetGroup

## VB6 Procedure

Function GetGroup(ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of groups selected for cold formed design.

MyName

This is an array that includes the name of each group selected for cold formed design.

## Remarks

This function retrieves the names of all groups selected for cold formed design.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColdFormedDesignGroup()

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'select group for cold formed design

      ret = SapModel.DesignColdFormed.SetGroup("ALL", True)



   'get groups selected for cold formed design

      ret = SapModel.DesignColdFormed.GetGroup(NumberItems, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetGroup](SetGroup_%7BCold_Formed%7D.md)
