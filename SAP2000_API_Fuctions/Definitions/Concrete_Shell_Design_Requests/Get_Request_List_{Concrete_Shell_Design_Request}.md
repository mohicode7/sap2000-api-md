# GetRequestList {Concrete Shell Design Request}

## Syntax

SapObject.SapModel.DesignConcreteShell.DesignRequest.GetRequestList

## VB6 Procedure

Function GetRequestList (ByRef NumberRequest As Integer, ByRef RequestList() As String) As Long

## Parameters

NumberRequest

The number of design requests existing in the model.

RequestList

This is an array containing the name of all the design requests existing in the model.

## Remarks

This function retrieves the list of design requests in the model.

The function returns zero if the list of design request is successfully retrieved; otherwise it returns a nonzero value.

The function is not applicable for the Eurocode 2-2004 concrete shell design code.

## VBA Example



```
Sub GetDesignRequestList ()

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
  

'add new load pattern
ret = SapModel.LoadPatterns.Add("Fluid", eLoadPatternType.EarthHydrostatic)
  

'add point load
Dim LoadValue() As Double
ReDim LoadValue(5)
LoadValue(1) = 10
  
ret = SapModel.PointObj.SetLoadForce("14", "DEAD", LoadValue)
  

'add combo
ret = SapModel.RespCombo.Add("COMB1", 1)
ret = SapModel.RespCombo.Add("COMB2", 1)
ret = SapModel.RespCombo.Add("COMB3", 1)
ret = SapModel.RespCombo.Add("COMB4", 1)
  

'add load case to combo
ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType.LoadCase, "DEAD", 1.1)
ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType.LoadCase, "DEAD", 1.2)
ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType.LoadCase, "DEAD", 1.3)
ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType.LoadCase, "DEAD", 1.4)
  

'set concrete shell design code
ret = SapModel.DesignConcreteShell.SetCode("ACI 350-20")
  

Dim RequestName As String
Dim GroupList() As String
Dim ComboList() As String
Dim AutoCombo As Boolean
Dim AutoComboCaseList() As String
Dim NewName As String
Dim NumberRequest As Integer
Dim RequestList() As String
Dim NumberGroup As Integer
Dim NumberCombo As Integer
  

RequestName = "R1"
ReDim RequestList(1)
RequestList(0) = ""
RequestList(1) = ""
ReDim GroupList(1)
ReDim ComboList(1)
ReDim AutoComboCaseList(2)
GroupList(1) = "GROUP1"
ComboList(1) = "COMB1"
AutoCombo = True
AutoComboCaseList(1) = "DEAD"
AutoComboCaseList(2) = "Fluid"
ret = SapModel.DesignConcreteShell.DesignRequest.Add(RequestName, GroupList, ComboList, AutoCombo, AutoComboCaseList)
  

NewName = "R11"
ret = SapModel.DesignConcreteShell.DesignRequest.ChangeName(RequestName, NewName)
  

RequestName = "R2"
GroupList(1) = "GROUP2"
ComboList(1) = "COMB2"
AutoCombo = False
ret = SapModel.DesignConcreteShell.DesignRequest.Add(RequestName, GroupList, ComboList, AutoCombo, AutoComboCaseList)


RequestName = "R123"
ReDim GroupList(3)
ReDim ComboList(3)
GroupList(1) = "GROUP1"
GroupList(2) = "GROUP2"
GroupList(3) = "GROUP3"
ComboList(1) = "COMB1"
ComboList(2) = "COMB2"
ComboList(3) = "COMB3"
  
ret = SapModel.DesignConcreteShell.DesignRequest.Add(RequestName, GroupList, ComboList, AutoCombo, AutoComboCaseList)

ret = SapModel.DesignConcreteShell.DesignRequest.GetRequestList(NumberRequest, RequestList)

      

       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
```



## Release Notes

Initial release in version 26.0.0

## See Also
