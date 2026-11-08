# UnsetAsActiveObject

## Syntax

SapObject.UnsetAsActiveObject

## VB6 Procedure

Function UnsetAsActiveObject () As Long

## Parameters

None

## Remarks

This function removes the current instance of a SapObject from the system Running Object Table (ROT).

This function returns zero if the current instance is successfully removed from the system ROT and nonzero if it fails or if the current instance is not in the system ROT.

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



   'unset as active SapObject instance

      ret = SapObject.UnsetAsActiveObject



   'close Sap2000

      SapObject.ApplicationExit False

   Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.01.

## See Also

[SetAsActiveObject](SetAsActiveObject.md)
