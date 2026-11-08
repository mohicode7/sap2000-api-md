# ChangeName

## Syntax

SapObject.SapModel.NamedAssign.ReleaseFrame.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined frame end release.

NewName

The new name for the frame end release.

## Remarks

This function changes the name of an existing frame end release.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeFrameEndReleaseName()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ii() As Boolean

      Dim jj() As Boolean

      Dim StartValue() As Double

      Dim EndValue() As Double



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



   'add new end release

      ReDim ii(5)

      ReDim jj(5)

      ReDim StartValue(5)

      ReDim EndValue(5)

      ii(0) = True

      StartValue(0) = 10

      ret = SapModel.NamedAssign.ReleaseFrame.SetReleases("FREL1", ii, jj, StartValue, EndValue)



   'change name of end release

      ret = SapModel.NamedAssign.ReleaseFrame.ChangeName("FREL1", "MyRelease")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
