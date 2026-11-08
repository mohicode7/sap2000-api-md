# SetFromFile

## Syntax

SapObject.SapModel.Func.FuncPSD.SetFromFile

## VB6 Procedure

Function SetFromFile(ByVal Name As String, ByVal FileName As String, ByVal HeadLines As Long, ByVal PreChars As Long, ByVal PointsPerLine As Long, ByVal ValueType As Long, ByVal FreeFormat As Boolean, Optional ByVal NumberFixed As Long = 10, Optional ByVal FreqTypeInFile As Long = 1) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

FileName

The full path of the text file containing the function data.

HeadLines

The number of header lines in the text file to be skipped before starting to read function data.

PreChars

The number of prefix characters to be skipped on each line in the text file.

PointsPerLine

The number of function points included on each text file line.

ValueType

This is either 1 or 2, indicating value type.

1 = Values at equal time intervals

2 = Time and function values

FreeFormat

This item is True if the data is provided in a free format. It is False if it is in a fixed format.

NumberFixed

This item only applies when the FreeFormat item is False. It is the number of characters per item.

FreqTypeInFile

This is either 1 or 2, indicating frequency type.

1 = Hz

2 = RPM

## Remarks

This function defines a power spectral density function from file.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetPSDFuncFromFile()

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



   'add PSD function from file

      ret = SapModel.Func.FuncPSD.SetFromFile("PSD-1", "C:\SapAPI\FuncPSD.txt", 2, 0, 1, 2, True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Text File

Following is the contents of the text file name FuncPSD.txt used in the VBA Example.



```
Power Spectral Density Function
One pair of Frequency (Hz) and Value items per line
0  1
1  1
2  2
3  2
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetFromFile](GetFromFile_%7BPSD%7D.md)
