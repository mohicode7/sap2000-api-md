# VerifyPassed

## Syntax

SapObject.SapModel.DesignConcrete.VerifyPassed

## VB6 Procedure

Function VerifyPassed(ByRef NumberItems As Long, ByRef n1 As Long, ByRef n2 As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of concrete frame objects that did not pass the design check or have not yet been checked.

n1

The number of concrete frame objects that did not pass the design check.

n2

The number of concrete frame objects that have not yet been checked.

MyName

This is an array that includes the name of each frame object that did not pass the design check or has not yet been checked.

## Remarks

This function retrieves the names of the frame objects that did not pass the design check or have not yet been checked, if any.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub VerifyConcreteDesignPassed()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyName2() As String

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'verify frame objects successfully designed

      ret = SapModel.DesignConcrete.VerifyPassed(NumberItems, n1, n2, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
