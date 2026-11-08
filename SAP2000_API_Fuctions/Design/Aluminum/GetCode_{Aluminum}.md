# GetCode

## Syntax

SapObject.SapModel.DesignAluminum.GetCode

## VB6 Procedure

Function GetCode(ByRef CodeName As String) As Long

## Parameters

CodeName

This is one of the following aluminum design code names.

AA-ASD 2000

AA-LRFD 2000

## Remarks

This function retrieves the aluminum design code.

The function returns zero if the code is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignCode()

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



   'add aluminum material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)



   'create new aluminum frame section property

      ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



   'get aluminum design code

      ret = SapModel.DesignAluminum.GetCode(CodeName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetCode](SetCode_%7BAluminum%7D.md)
