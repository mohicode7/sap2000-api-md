# SetAsActiveObject

## Syntax

SapObject.SetAsActiveObject

## VB6 Procedure

Function SetAsActiveObject () As Long

## Parameters

None

## Remarks

This function sets the active instance of a SapObject in the system Running Object Table (ROT), replacing the previous instance(s).

When a new SapObject is created using the API, it is automatically added to the system ROT if none is already present. Subsequent instances of the SapObject do not alter the ROT as long as at least one active instance of a SapObject is present in the ROT.

The Windows API call GetObject() can be used to attach to the active SapObject instance registered in the ROT.

This function returns zero if the current instance is successfully added to the system ROT and nonzero if it fails.

## VBA Example



```
Sub StartExample()

   'dimension variables

      Dim SapObject as cOAPI

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      ret = SapObject.ApplicationStart



   'set as active SapObject instance

      ret = SapObject.SetAsActiveObject



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.01.

## See Also

[UnsetAsActiveObject](UnsetAsActiveObject.md)
