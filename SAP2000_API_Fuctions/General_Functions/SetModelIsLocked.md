# SetModelIsLocked

## Syntax

SapObject.SapModel.SetModelIsLocked

## VB6 Procedure

Function SetModelIsLocked(LockIt as Boolean) As Long

## Parameters

LockIt

The item is True if the model is to be locked and False if it is to be unlocked.

## Remarks

The function returns zero if the locked status of the model is successfully set. Otherwise it returns a nonzero value.

With some exceptions, definitions and assignments can not be changed in a model while the model is locked. If an attempt is made to change a definition or assignment while the model is locked and that change is not allowed in a locked model, an error will be returned.

## VBA Example



```
Sub SetModelLocked()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'Lock model

      ret = SapModel.SetModelIsLocked(True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetModelIsLocked](GetModelIsLocked.md)
