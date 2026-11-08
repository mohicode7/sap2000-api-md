# SetNone

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeastate.SetNone

## VB6 Procedure

Function SetNone(ByVal Name As String) As Long

## Parameters

Name

The name of an existing seastate-type load pattern.

## Remarks

This function sets the auto seastate loading type for the specified load pattern to None.

The function returns zero if the loading type is successfully assigned; otherwise it returns a nonzero vale.

## VBA Example



```
Sub AssignSeastateNone()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim MyCenter(2) As Double
      Dim MyParams(8) As Double
      

   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)
 
   'define diaphragm constraints
      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", Z)
      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph2", Z)



   'assign points to diaphragm
      ret = SaModel.SelectObj.ClearSelection
      ret = SapModel.SelectObj.PlaneXY("2)
      ret = SapModel.PointObj.SetConstraint("", "Diaph1", SelectedObjects)
      ret = SapModel.SelectObj.ClearsSelection
      ret = SapModel.SelectObj.PlaneXY("3")
      ret = SapModel.PointObject.SetConstraint("", Diaph2", SelectedObjects)
      ret = SapModel.SelectObj.ClearSelection
   
   'add new load pattern

      ret = SapModel.LoadPatterns.Add("SEASTATE", LTYPE_SEASTATE)
 
   'assign seastate parameters
      MyCenter(0) = 0
      MyCenter(1) = 0
      MyCenter(2) = 0
      MyParams(0) = 0
      MyParams(1) = 0
      MyParams(2) = 0
      MyParams(3) = 0.5
      MyParams(4) = 0
      MyParams(5) = 0
      MyParams(6) = 0
      MyParams(7) = 0
      MyParams(8) = 0
      ret = SapModel.LoadPatterns.AutoSeastate.SetAuto("SEASTATE", 3, "Global", True, MyCenter, MyParms)
 
   'set auto seastate loading type to None
      ret = SapModel.LoadPatterns.AutoSWind.SetNone("SEASTATE")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

## See Also
