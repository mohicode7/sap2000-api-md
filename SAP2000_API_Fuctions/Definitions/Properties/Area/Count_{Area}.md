# Count

## Syntax

SapObject.SapModel.PropArea.Count

## VB6 Procedure

Function Count(Optional ByVal PropType As Long = 0) As Long

## Parameters

PropType

This optional value is 0, 1, 2 or 3, indicating the type of area properties included in the count.

0 = All

1 = Shell

2 = Plane

3 = Asolid

## Remarks

This function returns the total number of defined area properties in the model. If desired, counts can be returned for all area properties of a specified type in the model.

## VBA Example



```
Sub CountAreaProps()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'return number of defined area properties of all types

      Count = SapModel.PropArea.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
