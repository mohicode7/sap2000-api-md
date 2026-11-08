# Count

## Syntax

SapObject.SapModel.LinkObj.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns a count of the link objects in the model.

## VBA Example



```
Sub CountLinkObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim Name1 As String

      Dim Name2 As String



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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("3", "", Name1, True)

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name2)



   'refresh view

      ret = SapModel.View.RefreshView



   'get number of link objects

      Count = SapModel.LinkObj.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
