# GetModelIsLocked

## Syntax

SapObject.SapModel.GetModelIsLocked

## VB6 Procedure

Function GetModelIsLocked() As Boolean

## Parameters

None

## Remarks

The function returns True if the model is locked and False if it is unlocked.

With some exceptions, definitions and assignments can not be changed in a model while the model is locked. If an attempt is made to change a definition or assignment while the model is locked and that change is not allowed in a locked model, an error will be returned.

## VBA Example



```
Sub GetModelLocked()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim IsLocked As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 124, 3, 200)



   'check if model is locked

      IsLocked = SapModel.GetModelIsLocked



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetModelIsLocked](SetModelIsLocked.md)
