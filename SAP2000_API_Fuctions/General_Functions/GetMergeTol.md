# GetMergeTol

## Syntax

SapObject.SapModel.GetMergeTol

## VB6 Procedure

Function GetMergeTol(ByRef MergeTol As Double) As Long

## Parameters

MergeTol

The program auto merge tolerance. [L]

## Remarks

This function retrieves the value of the program auto merge tolerance.

The function returns zero if the tolerance is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAutoMergeTolerance()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MergeTol As Double



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



   'get auto merge tolerance

      ret = SapModel.GetMergeTol(MergeTol)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetMergeTol](SetMergeTol.md)
