# RefreshWindow

## Syntax

SapObject.SapModel.View.RefreshWindow

## VB6 Procedure

Function RefreshWindow(Optional ByVal Window As Long = 0) As Long

## Parameters

Window

This is 0 meaning all windows or an existing window number. It indicates the window(s) to be refreshed.

## Remarks

This function refreshes the specified window(s). It returns zero if the windows are successfully refreshed, otherwise it returns a nonzero value.

See [RefreshWindow and RefreshView](RefreshWindow_and_RefreshView.md) for more information.

## VBA Example



```
Sub RefreshAllWindows()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Boolean



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



   'add joint restraint

      Redim Value(5)

      For i = 0 to 5

         Value(i) = True

      Next i

      ret = SapModel.PointObj.SetRestraint("1", Value)



   'refresh all windows

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[RefreshView](RefreshView.md)
