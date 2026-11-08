# GetDesignSection

## Syntax

SapObject.SapModel.DesignAluminum.GetDesignSection

## VB6 Procedure

Function GetDesignSection(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a frame object with a aluminum frame design procedure.

PropName

The name of the design section for the specified frame object.

## Remarks

This function retrieves the design section for a specified aluminum frame object.

The function returns zero if the section is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignSection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim PropName As String



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



   'get design section

      ret = SapModel.DesignAluminum.GetDesignSection("8", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetDesignSection](SetDesignSection_%7BAluminum%7D.md)
