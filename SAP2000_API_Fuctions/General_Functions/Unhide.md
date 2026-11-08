# Unhide

## Syntax

SapObject.Unhide

## VB6 Procedure

Function Unhide() As Long

## Parameters

None

## Remarks

This function unhides the Sap2000 application, that is, it makes it visible. When the application is hidden, it is not visible on the screen or on the Windows task bar.

The function returns zero if the Sap2000 application is successfully unhidden (set visible) and nonzero if the function fails. If the application is already visible (not hidden) calling this function returns an error.

## VBA Example



```
Sub UnhideSap()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'create model from template

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 124, 3, 200)



   'hide application

      ret = SapObject.Hide



   'unhide application

      ret = SapObject.Unhide



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Visible](Visible.md)

[Hide](Hide.md)
