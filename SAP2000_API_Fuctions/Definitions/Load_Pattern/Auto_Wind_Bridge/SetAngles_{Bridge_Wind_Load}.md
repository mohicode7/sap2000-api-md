# SetAngles {Bridge Wind Load}

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetAngles

## VB6 Procedure

Function SetAngles(ByVal Name As String, ByVal SymmetricTran As Boolean, ByVal SymmetricLong As Boolean, ByVal UserDefined As Boolean, ByVal NumberAngles As Long, ByRef Angles() As Double, ByRef TranCoeff() As Double, ByRef LongCoeff() As Double) As Long

## Parameters

Name

The name of an existing bridge wind load pattern.

SymmetricTran

This is True if the angles are symmetric about the transverse axis, otherwise False.

SymmetricLong

This is True if the angles are symmetric about the longitudinal axis, otherwise False.

UserDefined

This is True if the angles are user defined, otherwise False.

NumberAngles

The number of angles specified for the load pattern. This is only used if UserDefined = True.

Angles

This is an array containing the wind angles. If duplicates are included, only the first one will be used. If SymmetricTran = True and/or SymmetricLong = True, only angles between 0 and 90 degrees will be used. Additional angles will be generated via the symmetry conditions. This is only used if UserDefined = True.

TranCoeff

This is an array containing the transverse coefficients corresponding with the angles. This is only used if UserDefined = True.

LongCoeff

This is an array containing the longitudinal coefficients corresponding with the angles. This is only used if UserDefined = True.

## Remarks

This function sets the bridge wind angle data for a specified load pattern.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAngles()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim Angles() As Double
      Dim TranCoeff() As Double
      Dim LongCoeff() As Double
 
   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel
 
   'open existing model containing a bridge object

      ret = SapModel.File.OpenFile(“C:\Temp\BridgeModel.bdb” )



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WIND)



   'assign AASHTO2018 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge.SetAASHTO2018("WIND", 1, True, True, True, -100, True, 10, True, 10, 100, 2, 1, 1.3, 1.6)
 
   'set angle data
      ReDim Angles(4)
      ReDim TranCoeff(4)
      ReDim LongCoeff(4)
      Angles(0) = 0
      TranCoeff(0) = 1
      LongCoeff(0) = 0
      Angles(1) = 20
      TranCoeff(1) = 0.85
      LongCoeff(1) = 0.18
      Angles(2) = 40
      TranCoeff(2) = 0.7
      LongCoeff(2) = 0.26
      Angles(3) = 60
      TranCoeff(3) = 0.34
      LongCoeff(3) = 0.38
      Angles(4) = 80
      TranCoeff(4) = 0.12
      LongCoeff(4) = 0.42
      ret = SapModel.LoadPatterns.AutoWindBridge.SetAngles("WIND", False, False, True, 5, Angles, TranCoeff, LongCoeff)
   

   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in v21.0.0.

## See Also

[GetAngles](GetAngles_%7BBridge_Wind_Load%7D.md)
