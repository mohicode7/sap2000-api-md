# Hide

## Syntax

SapObject.Hide

## VB6 Procedure

Function Hide() As Long

## Parameters

None

## Remarks

This function hides the Sap2000 application. When the application is hidden it is not visible on the screen or on the Windows task bar.

The function returns zero if the Sap2000 application is successfully hidden and nonzero if the function fails. If the application is already hidden calling this function returns an error.

## VBA Example



```
Sub HideSap()

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

[Unhide](Unhide.md)
