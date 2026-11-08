# CountLoadDispl

## Syntax

SapObject.SapModel.PointObj.CountLoadDispl

## VB6 Procedure

Function CountLoadDispl(ByRef Count As Long, Optional ByVal Name As String = "", Optional ByVal LoadPat As String = "") As Long

## Parameters

Count

The number of counted ground displacement loads.

Name

This optional item is the name of an existing point object.

LoadPat

This optional item is the name of an existing load pattern.

## Remarks

If neither the Name item nor the LoadPat item is provided, the Count item returns the total number of ground displacement load assignments in the model.

If the Name item is provided but not the LoadPat item, the Count item returns the total number of ground displacement load assignments made for the specified point object.

If the Name item is not provided but the LoadPat item is specified, the Count item returns the total number of ground displacement load assignments made to all point objects for the specified load pattern.

If both the Name item and the LoadPat item are provided, the Count item returns the total number of ground displacement load assignments made to the specified point object for the specified load pattern.

If the Name item or the LoadPat item is provided but is not recognized by the program as valid, an error is returned.

This function returns zero if the count is successfully completed, otherwise it returns a nonzero value.

## VBA Example



```
Sub CountGroundDisplacementLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value() As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add ground displacement load

      Redim Value(5)

      Value(0) = 10

      ret = SapModel.PointObj.SetLoadDispl("1", "DEAD", Value)



   'get number of ground displacement loads

      ret = SapModel.PointObj.CountLoadDispl(Count)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoadDispl](GetLoadDispl_%7BPoint_Object%7D.md)

[SetLoadDispl](SetLoadDispl.md)

[DeleteLoadDispl](DeleteLoadDispl.md)
