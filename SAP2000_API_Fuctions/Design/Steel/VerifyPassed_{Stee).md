# VerifyPassed

## Syntax

SapObject.SapModel.DesignSteel.VerifyPassed

## VB6 Procedure

Function VerifyPassed(ByRef NumberItems As Long, ByRef n1 As Long, ByRef n2 As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of steel frame objects that did not pass the design check or have not yet been checked.

n1

The number of steel frame objects that did not pass the design check.

n2

The number of steel frame objects that have not yet been checked.

MyName

This is an array that includes the name of each frame object that did not pass the design check or has not yet been checked.

## Remarks

This function retrieves the names of the frame objects that did not pass the design check or have not yet been checked, if any.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub VerifySteelDesignPassed()

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



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'import new frame section properties

          ret = SapModel.PropFrame.ImportProp("W18X35", "A992Fy50", "AISC16.XML", "W18X35")

          ret = SapModel.PropFrame.ImportProp("W18X40", "A992Fy50", "AISC16.XML", "W18X40")

          ret = SapModel.PropFrame.ImportProp("W18X46", "A992Fy50", "AISC16.XML", "W18X46")



       'define new auto select list frame section property

          ReDim MyName2(2)

          MyName2(0) = "W18X35"

          MyName2(1) = "W18X40"

          MyName2(2) = "W18X46"

          ret = SapModel.PropFrame.SetAutoSelectSteel("AUTO1", 3, MyName2)



       'set frame section properties

          ret = SapModel.FrameObj.SetSection("8", "AUTO1")

          ret = SapModel.FrameObj.SetSection("10", "AUTO1")



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'start steel design

          ret = SapModel.DesignSteel.StartDesign



       'verify frame objects successfully designed

          ret = SapModel.DesignSteel.VerifyPassed(NumberItems, n1, n2, MyName)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.03.

## See Also
