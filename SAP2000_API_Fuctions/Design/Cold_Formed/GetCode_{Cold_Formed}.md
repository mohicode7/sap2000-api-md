# GetCode

## Syntax

SapObject.SapModel.DesignColdFormed.GetCode

## VB6 Procedure

Function GetCode(ByRef CodeName As String) As Long

## Parameters

CodeName

This is one of the following cold formed design code
names.

AISI-ASD96

AISI-LRFD96

## Remarks

This function retrieves the cold formed design code.

The function returns zero if the code is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColdFormedDesignCode()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim CodeName As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, 
 MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", 
 Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 72, 2, 144, True, "CdC", "CdC")



   'get cold formed design code

      ret = SapModel.DesignColdFormed.GetCode(CodeName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetCode](SetCode_%7BCold_Formed%7D.md)
