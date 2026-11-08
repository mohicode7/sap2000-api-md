# GetFromFile

## Syntax

SapObject.SapModel.Func.FuncSS.GetFromFile

## VB6 Procedure

Function GetFromFile(ByVal Name As String, ByRef FileName As String, ByRef HeadLines As Long, ByRef PreChars As Long, ByRef PointsPerLine As Long, ByRef ValueType As Long, ByRef FreeFormat As Boolean, ByRef NumberFixed As Long, ByRef FreqTypeInFile As Long) As Long

## Parameters

Name

The name of a defined steady state function specified to be from a text file.

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

This function retrieves the definition of a steady state function from file.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSSFuncFromFile()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim HeadLines As Long

      Dim PreChars As Long

      Dim PointsPerLine As Long

      Dim ValueType As Long

      Dim FreeFormat As Boolean

      Dim NumberFixed As Long

      Dim FreqTypeInFile As Long



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



   'add SS function from file

      ret = SapModel.Func.FuncSS.SetFromFile("SS-1", "C:\SapAPI\FuncSS.txt", 2, 0, 1, 2, True)



   'get SS function from file

      ret = SapModel.Func.FuncSS.GetFromFile("SS-1", FileName, HeadLines, PreChars, PointsPerLine, ValueType, FreeFormat, NumberFixed, FreqTypeInFile)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Text File

Following is the contents of the text file name FuncSS.txt used in the VBA Example.



```
Steady State Function
One pair of Frequency (Hz) and Value items per line
0  1
1  1
2  2
3  2
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetFromFile](SetFromFile_%7BSteady_State%7D.md)
