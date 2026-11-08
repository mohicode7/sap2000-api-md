# ConvertLineToBLL

## Syntax

SapObject.SapModel.GenRefLine.ConvertLineToBLL

## VB6 Procedure

Function ConvertLineToBLL(ByVal Name As String, Optional ByVal FirstStation As Double = 0, Optional ByVal CSys As String = "Global", Optional ByVal OffsetX As Double = 0, Optional ByVal OffsetY As Double = 0, Optional ByVal OffsetZ As Double) As Long

## Parameters

Name

This is the name of an existing general reference line.

FirstStation

The first station value on the bridge layout line. [L]

CSys

The name of the coordinate system in which the general reference line is offset to create the bridge layout line.

OffsetX

The distance to offset the general reference line in the x-direction of the specified CSys, to the location of the new bridge layout line.

OffsetY  
The distance to offset the general reference line in the y-direction of the specified CSys, to the location of the new bridge layout line.

OffsetZ  
The distance to offset the general reference line in the z-direction of the specified CSys, to the location of the new bridge layout line.

## Remarks

This function converts an existing general reference line to a new bridge layout line, with the ability to specify offset values.

The function returns zero if the bridge layout line is successfully created; otherwise it returns a nonzero value.

## VBA Example



```
Sub ConvertGenRefLine()

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



   'convert general reference line to bridge layout line

      ret = SapModel.GenRefLine.ConvertLineToBLL("GRef1")



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
