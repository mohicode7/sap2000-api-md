# Delete

## Syntax

SapObject.SapModel.LoadPatterns.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing load pattern.

## Remarks

The function deletes the specified load pattern.

The function returns zero if the load pattern is successfully deleted, otherwise it returns a nonzero value.

The load pattern is not deleted and the function returns an error if the load pattern is assigned to an load case or if it is the only defined load pattern.

## VBA Example



```
Sub DeleteLoadPattern()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = New Sap2000v16.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'delete load case

   'Note:  ret in the above line returns 1, indicating an error.

   'This is the case because the load pattern DEAD is assigned to

   'load case DEAD. This is also the case because the load

   'case DEAD is the only load pattern defined.

      ret = SapModel.LoadPatterns.Delete("DEAD")



   'Note:  ret in the above line returns 1, indicating an error. This is the case because the load pattern DEAD is assigned to load case DEAD. This is also the case because the load pattern DEAD is the only load pattern defined.



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
