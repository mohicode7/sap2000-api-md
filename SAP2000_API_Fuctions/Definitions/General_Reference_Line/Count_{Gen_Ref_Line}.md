# Count

## Syntax

SapObject.SapModel.GenRefLine.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None.

## Remarks

The function returns the number of defined general reference lines.

## VBA Example



```
Sub CountGenRefLine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'define new general reference line

      ret = SapModel.GenRefLine.SetLine("GRef1", 120, 5)



   'get number of general reference line

      ret = SapModel.GenRefLine.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.0.

## See Also

[SetLine](SetLine_%7BGen_Ref_Line%7D.md)
