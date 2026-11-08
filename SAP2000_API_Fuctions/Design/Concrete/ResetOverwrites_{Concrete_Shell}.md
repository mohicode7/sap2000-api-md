# ResetOverwrites {Concrete Shell}

## Syntax

SapObject.SapModel.DesignConcreteShell.ResetOverwrites

## VB6 Procedure

Function ResetOverwrites() As Long

## Parameters

None

## Remarks

This function resets all concrete shell design overwrites to default values.

The function returns zero if the overwrites are successfully reset; otherwise it returns a nonzero value.

The function will fail if no concrete shell objects are present.

The function is not applicable for the Eurocode 2-2004 concrete shell design code.

## VBA Example



```
Sub ResetConcreteShellDesignOverwrites()
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
ret = SapModel.InitializeNewModel(eUnits.kip_in_F) 
  

'create a wall model from template
ret = SapModel.File.NewWall(6, 4, 6, 4)
  

'initialize new material property
ret = SapModel.PropMaterial.SetMaterial("Rebar", eMatType.Rebar)
  

'assign other properties
ret = SapModel.PropMaterial.SetORebar_1("Rebar", 62, 93, 70, 102, 2, 2, 0.02, 0.1, -0.1, False)
  

'set new area property
ret = SapModel.PropArea.SetShell_1("A1", 1, True, "4000Psi", 0, 16, 16)
 

'set area property design parameters
ret = SapModel.PropArea.SetShellDesign("A1", "Rebar", 2, 2, 3, 2.5, 3.5)
   

'set area property
ret = SapModel.AreaObj.SetProperty("4", "A1")
   

'define new group
ret = SapModel.GroupDef.SetGroup("GROUP1")
ret = SapModel.GroupDef.SetGroup("GROUP2")
ret = SapModel.GroupDef.SetGroup("GROUP3")
ret = SapModel.GroupDef.SetGroup("GROUP4")
  

ret = SapModel.AreaObj.SetGroupAssign("1", "GROUP1")
ret = SapModel.AreaObj.SetGroupAssign("2", "GROUP2")
ret = SapModel.AreaObj.SetGroupAssign("3", "GROUP3")
ret = SapModel.AreaObj.SetGroupAssign("4", "GROUP4")
   

'set concrete shell design code
ret = SapModel.DesignConcreteShell.SetCode("ACI 350-20")
   

Dim value As Double
Dim progdet As Boolean
ret = SapModel.DesignConcreteShell.ACI350_20.SetPreference(5, 0.7)
ret = SapModel.DesignConcreteShell.ACI350_20.GetPreference(5, value)
 

ret = SapModel.DesignConcreteShell.ACI350_20.SetOverwrite("1", 4, 1.345)
ret = SapModel.DesignConcreteShell.ACI350_20.GetOverwrite("1", 4, value, progdet)
   
ret = SapModel.DesignConcreteShell.ResetOverwrites

            

           'close Sap2000

              SapObject.ApplicationExit False

              Set SapModel = Nothing

              Set SapObject = Nothing

    End Sub
```



## Release Notes

Initial release in version 26.0.0

## See Also
