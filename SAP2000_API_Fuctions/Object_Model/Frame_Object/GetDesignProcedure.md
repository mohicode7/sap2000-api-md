# GetDesignProcedure

## Syntax

SapObject.SapModel.FrameObj.GetDesignProcedure

## VB6 Procedure

Function GetDesignProcedure(ByVal Name As String, ByRef MyType As Long) As Long

## Parameters

Name

The name of an existing frame object.

MyType

This is 1, 2, 7, 8 or 9, indicating the design procedure for the specified frame object.

1 = Steel

2 = Concrete

7 = Aluminum

8 = Cold Formed

9 = No Design

## Remarks

This function retrieves the design procedure for a frame object.

The function returns zero if the design procedure is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameObjDesignProcedure()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long



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



   'get design procedure

      ret = SapModel.FrameObj.GetDesignProcedure("8", MyType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetDesignProcedure](SetDesignProcedure.md)
