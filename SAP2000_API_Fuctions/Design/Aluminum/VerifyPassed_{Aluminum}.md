# VerifyPassed

## Syntax

SapObject.SapModel.DesignAluminum.VerifyPassed

## VB6 Procedure

Function VerifyPassed(ByRef NumberItems As Long, ByRef n1 As Long, ByRef n2 As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of aluminum frame objects that did not pass the design check or have not yet been checked.

n1

The number of aluminum frame objects that did not pass the design check.

n2

The number of aluminum frame objects that have not yet been checked.

MyName

This is an array that includes the name of each frame object that did not pass the design check or has not yet been checked.

## Remarks

This function retrieves the names of the frame objects that did not pass the design check or have not yet been checked, if any.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub VerifyAluminumDesignPassed()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim n1 As Long

      Dim n2 As Long

      Dim MyName() As String



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



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start aluminum design

      ret = SapModel.DesignAluminum.StartDesign



   'verify frame objects successfully designed

      ret = SapModel.DesignAluminum.VerifyPassed(NumberItems, n1, n2, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
