# GetDesignSection

## Syntax

SapObject.SapModel.DesignSteel.GetDesignSection

## VB6 Procedure

Function GetDesignSection(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a frame object with a steel frame design procedure.

PropName

The name of the design section for the specified frame object.

## Remarks

This function retrieves the design section for a specified steel frame object.

The function returns zero if the section is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignSection()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim MyName() As String

          Dim PropName As String



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

          ReDim MyName(2)

          MyName(0) = "W18X35"

          MyName(1) = "W18X40"

          MyName(2) = "W18X46"

          ret = SapModel.PropFrame.SetAutoSelectSteel("AUTO1", 3, MyName)



       'set frame section properties

          ret = SapModel.FrameObj.SetSection("8", "AUTO1")

          ret = SapModel.FrameObj.SetSection("10", "AUTO1")



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'start steel design

          ret = SapModel.DesignSteel.StartDesign



       'get design section

          ret = SapModel.DesignSteel.GetDesignSection("8", PropName)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetDesignSection](SetDesignSection_%7BSteel%7D.md)
