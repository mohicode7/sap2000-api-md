# GetCode

## Syntax

SapObject.SapModel.DesignConcrete.GetCode

## VB6 Procedure

Function GetCode(ByRef CodeName As String) As Long

## Parameters

CodeName

This is one of the following concrete design code names.

AASHTO LRFD 2014  
AASHTO LRFD 2012  
AASHTO Concrete 07  
ACI 318-14  
ACI 318-11  
ACI 318-08/IBC2009  
AS 3600-09  
BS8110 97  
Chinese 2010  
CSA A23.3-14  
CSA A23.3-04  
Eurocode 2-2004  
Hong Kong CP 2013  
Indian IS 456-2000  
Italian NTC 2008  
KBC 2009  
Mexican RCDF 2004  
NZS 3101:2006  
Singapore CP 65:99  
SP 63.13330.2012  
TS 500-2000

## Remarks

This function retrieves the concrete design code.

The function returns zero if the code is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignCode()

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'get concrete design code

      ret = SapModel.DesignConcrete.GetCode(CodeName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Modified list of codes contained in the CodeName Parameter in version 15.0.1.

Updated list of available codes in v17.3.0.

Removed older codes which have been removed from the program in v18.0.0.

Updated list of available codes in v19.1.0.

## See Also

[SetCode](SetCode_%7BConcrete%7D.md)
