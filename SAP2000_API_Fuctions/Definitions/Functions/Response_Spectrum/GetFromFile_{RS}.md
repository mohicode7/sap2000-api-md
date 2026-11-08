# GetFromFile

## Syntax

SapObject.SapModel.Func.FuncRS.GetFromFile

## VB6 Procedure

Function GetFromFile(ByVal Name As String, ByRef FileName As String, ByRef HeadLines As Long, ByRef DampRatio As Double, ByRef ValueType As Long) As Long

## Parameters

Name

The name of a defined response spectrum function specified to be from a text file.

FileName

The full path of the text file containing the function data.

HeadLines

The number of header lines in the text file to be skipped before starting to read function data.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

ValueType

This is either 1 or 2, indicating time value type.

1 = Frequency

2 = Period

## Remarks

This function retrieves the definition of a response spectrum function from file.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncFromFile()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim HeadLines As Long

      Dim DampRatio As Double

      Dim ValueType As Long



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



   'add RS function from file

      ret = SapModel.Func.FuncRS.SetFromFile("RS-1", "C:\SapAPI\FuncRS.txt", 3, 0.04)



   'get RS function from file

      ret = SapModel.Func.FuncRS.GetFromFile("RS-1", FileName, HeadLines, DampRatio, ValueType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Text File

Following is the contents of the text file name FuncRS.txt used in the VBA Example.



```
Reponse Spectrum Function
One pair of Period (sec) and Acceleration (g) values per line
Acceleration values at equal spacing of 0.01 seconds.
0.030  0.500
0.125  1.355
0.587  1.355
0.660  1.355
1.562  0.576
4.000  0.219
10.00  0.037
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetFromFile](SetFromFile_%7BRS%7D.md)
