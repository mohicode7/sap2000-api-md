# GetCommonTo

## Syntax

SapObject.SapModel.PointObj.GetCommonTo

## VB6 Procedure

Function GetCommonTo(ByVal Name As String, ByRef CommonTo As Long) As Long

## Parameters

Name

The name of a point object or a group depending on the value selected for ItemType item.

CommonTo

The total number of objects (line, area, solid and link) that connect to the specified point object.

## Remarks

This function returns the total number of objects (line, area, solid and link) that connect to the specified point object.

The function returns zero if the CommonTo is successfully calculated, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointCommonTo()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim CommonTo as long



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



   'get CommonTo

      ret = SapModel.PointObj.GetCommonTo("6", CommonTo)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
