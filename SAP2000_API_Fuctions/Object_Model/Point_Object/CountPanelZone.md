# CountPanelZone

## Syntax

SapObject.SapModel.PointObj.CountPanelZone

## VB6 Procedure

Function CountPanelZone() As Long

## Parameters

None

## Remarks

This function returns the total number of panel zone assignments to point objects in the model.

## VBA Example



```
Sub CountPanelZones()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim Count as Long

      Dim PropType as Long

      Dim Thickness As Double

      Dim Connectivity As Long

      Dim LocalAxisFrom As Long



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



   'add panel zone assignment

      PropType = 1 'elastic from column and doubler plate

      Thickness = 2

      Connectivity  = 0 'beams to other objects

      LocalAxisFrom = 0 'column

      ret = SapModel.PointObj.SetPanelZone("3", PropType, Thickness, 0, 0, "", Connectivity, LocalAxisFrom, 0)



   'get number of panel zone assignments

      Count = SapModel.PointObj.CountPanelZone



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPanelZone](GetPanelZone.md)

[SetPanelZone](SetPanelZone.md)

[DeletePanelZone](DeletePanelZone.md)
