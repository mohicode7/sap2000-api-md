# SetCode

## Syntax

SapObject.SapModel.DesignColdFormed.SetCode

## VB6 Procedure

Function SetCode(ByVal CodeName As String) As Long

## Parameters

CodeName

This is one of the following cold formed design code names.

AISI-ASD96

AISI-LRFD96

## Remarks

This function sets the cold formed design code.

The function returns zero if the code is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetColdFormedDesignCode()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144)



   'set cold formed design code

      ret = SapModel.DesignColdFormed.SetCode("AISI-LRFD96")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetCode](GetCode_%7BCold_Formed%7D.md)
