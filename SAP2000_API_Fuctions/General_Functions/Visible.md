# Visible

## Syntax

SapObject.Visible

## VB6 Procedure

Function Visible() As Boolean

## Parameters

None

## Remarks

The function returns True if the Sap2000 application is visible on the screen, otherwise it returns False.

## VBA Example



```
Sub IsVisible()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as long

      Dim Visible as Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'create model from template

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 124, 3, 200)



   'get application visibility

      Visible = SapObject.Visible



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Hide](Hide.md)

[Unhide](Unhide.md)
