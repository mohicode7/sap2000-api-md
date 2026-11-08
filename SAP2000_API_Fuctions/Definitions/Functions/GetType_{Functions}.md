# GetTypeOAPI

## Syntax

SapObject.SapModel.Func.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef FuncType As Long, ByRef AddType As Long) As Long

## Parameters

Name

The name of an existing function.

FuncType

This is one of the following numbers, indicating the type of function.

1 = Response spectrum

2 = Time history

3 = Power spectral density

4 = Steady state

AddType

The is one of the following items, indicating the function subtype.

Response Spectrum Functions

0 = From file

1 = User

2 = UBC 94

3 = UBC 97

4 = BOCA 96

5 = NBCC 95

6 = IBC 2003

7 = NEHRP 97

8 = Eurocode 8-1998

9 = NZS4203 1992

10 = Chinese 2010

11 = Italian Ordinanza 3274

12 = IS1893:2002

13 = AASHTO LRFD 2006

14 = NCHRP Project 20-07

15 = IBC 2006

16 = NBCC 2005  
17 = Eurocode 8-2004  
18 = AS 1170.4-2007  
19 = NZS 1170.5-2004  
20 = AASHTO 2007  
21 = Chinese JTG/T B02-2013  
22 = Chinese GB 50111-2006  
23 = IBC 2009  
24 = NBCC 2010  
25 = NTC 2008  
26 = AASHTO 2012  
27 = IBC 2012  
28 = TSC 2007  
29 = SI 413(1995)  
30 = Argentina INPRES-CIRSOC 103  
31 = Chile Norma NCh433+DS61  
32 = Chile Norma NCh2369-2003  
33 = Colombia NSR-10  
34 = Ecuador NEC-11 Capitulo 2  
35 = Guatemala AGIES NSE 2-10  
36 = Mexico NTC 2004  
37 = Peru Norma E.030  
38 = Dominican Republic R-001  
39 = Venezuela COVENIN 1756-2:2001  
40 = KBC 2009  
41 = Mexico CFE-93  
42 = Peru NTE E.030 2014  
43 = Mexico CFE-2008  
44 = Ecuado Norma NEC-SE-DS 2015  
45 = Costa Rica Seismic Code 2010  
46 = SP 14.13330.2014  
47 = Chinese CJJ 166-2011  
48 = NBCC 2015  
49 = IS 1893:2016  
50 = ASCE 7-16  
51 = KBC 2016  
52 = NTC 2018  
53 = TSC 2018

Time History Functions

0 = From file

1 = User

2 = Sine

3 = Cosine

4 = Ramp

5 = Sawtooth

6 = Triangular

7 = User periodic

Power Spectral Density Functions

0 = From file

1 = User

Steady State Functions

0 = From file

1 = User

## Remarks

This function retrieves the function type for the specified function.

The function returns zero if the type is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetFunctionType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FuncType As Long

      Dim AddType As Long



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



   'get function type

      ret = SapModel.Func.GetTypeOAPI("UNIFRS", FuncType, AddType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed function name to GetTypeOAPI in v17.0.0.

Updated list of AddType values for response spectrum functions in v18.2.0.

## See Also
