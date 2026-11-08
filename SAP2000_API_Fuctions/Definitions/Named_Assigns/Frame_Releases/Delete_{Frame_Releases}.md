# Delete

## Syntax

SapObject.SapModel.NamedAssign.ReleaseFrame.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing frame end release.

## Remarks

The function deletes a specified frame end release.

The function returns zero if the end release is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified end release can not be deleted; for example, if it is currently used by a staged construction load case.

## VBA Example



```
Sub DeleteFrameEndRelease()

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

      ii(0) = False

      StartValue(0) = 0

      ii(5) = True

      jj(5) = True

      ret = SapModel.NamedAssign.ReleaseFrame.SetReleases("FREL2", ii, jj, StartValue, EndValue)



   'delete end release

      ret = SapModel.NamedAssign.ReleaseFrame.Delete("FREL1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
