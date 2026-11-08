# SetCode

## Syntax

SapObject.SapModel.DesignSteel.SetCode

## VB6 Procedure

Function SetCode(ByVal CodeName As String) As Long

## Parameters

CodeName

This is one of the following steel design code names.

AASHTO LRFD 2007

AISC-ASD89

AISC 360-10

AISC360-05/IBC2006

AISC-LRFD93

API RP2A-LRFD 97

API RP2A-WSD2000

API RP2A-WSD2014

AS 4100-1998  
ASCE 10-97

BS5950 2000

Chinese 2010  
CSA S16-19  
CSA S16-14  
CSA-S16-09

EN 1993-1-1:2005 (formerly EUROCODE 3-2005)

Indian IS 800-2007

Italian NTC 2008

Italian UNI 10011

KBC 2009

Norsok N-004 2013

NZS 3404-1997

SP 16.13330.2011

## Remarks

This function sets the steel design code.

The function returns zero if the code is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignCode()

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

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'set steel design code

          ret = SapModel.DesignSteel.SetCode("AISC 360-10")



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.03.

In version 14.1.0, added Norway as a Country parameter.

Added Norsok N-004 2013 in version 16.1.0.

Updated list of code names in version 17.3.0.

Removed older codes which have been removed from the program in v18.0.0.

Updated list of available codes in v19.1.0.

## See Also

[GetCode](GetCode_%7BSteel%7D.md)
