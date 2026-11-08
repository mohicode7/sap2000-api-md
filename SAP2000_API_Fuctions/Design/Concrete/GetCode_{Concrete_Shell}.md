# GetCode (for Concrete Shell Design)

## Syntax

SapObject.SapModel.DesignConcreteShell.GetCode

## VB6 Procedure

Function GetCode(ByRef CodeName As String) As Long

## Parameters

CodeName

This is one of the following concrete shell design code names.

ACI 350-20

Eurocode 2-2004

## Remarks

This function retrieves the concrete shell design code.

The function returns zero if the code is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteShellDesignCode()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim CodeName As String



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create blank model

             ret = SapModel.File.NewBlank()



       'get concrete design code

          ret = SapModel.DesignConcreteShell.GetCode(CodeName)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 24.0.0

Updated to include ACI 350-20 in list of code names in version 26.0.0

## See Also

[SetCode {Concrete Shell}](SetCode_%7BConcrete_Shell%7D.md)
