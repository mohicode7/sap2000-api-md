# **SetFromFile\_1**

## Syntax

SapObject.SapModel.Func.FuncTH.SetFromFile\_1

## VB6 Procedure

Function SetFromFile\_1(ByVal Name As String, ByVal FileName As String, ByVal HeadLines As Long, ByVal PreChars As Long, ByVal PointsPerLine As Long, ByVal ValueType As Long, ByVal FreeFormat As Boolean, Optional ByVal NumberFixed As Long = 10, Optional ByVal DT As Double = 0.02) As Long

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

This is 1 or 2, indicating value type.

1 = Values at equal time intervals

2 = Time and function values

FreeFormat

This item is True if the data is provided in a free format. It is False if it is in a fixed format.

NumberFixed

This item applies only when the FreeFormat item is False. It is the number of characters per item.

DT

This item applies only when the ValueType item is 1 (equal time intervals). It is the time interval between function points.

## Remarks

This function defines a time history function from file.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncFromFile_1()

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



   'add TH function from file

      ret = SapModel.Func.FuncTH.SetFromFile_1("TH-1", "C:\SapAPI\FuncTH.txt", 3, 0, 3, 2, True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Text File

Following is the contents of the text file name FuncTH.txt used in the VBA Example.



```
Time History Function
Time (sec) and Acceleration (g) values
3 points per line
   0.00000    .01080    .04200    .00100    .09700    .01590

    .16100   -.00010    .22100    .01890    .26300    .00010

    .29100    .00590    .33200   -.00120    .37400    .02000

    .42900   -.02370    .47100    .00760    .58100    .04250

    .62300    .00940    .66500    .01380    .72000   -.00880

    .72010   -.02560    .78900   -.03870    .78910   -.05680

    .87200   -.02320    .87210   -.03430    .94100   -.04020

    .94110   -.06030    .99700   -.07890   1.06600   -.06660

   1.06610   -.03810   1.09400   -.04290   1.16800    .08970

   1.31500   -.16960   1.38400   -.08280   1.41200   -.08280

   1.44000   -.09450   1.48100   -.08850   1.50900   -.10800

   1.53700   -.12800   1.62800    .11440   1.70300    .23550

   1.80000    .14280   1.85500    .17770   1.92400   -.26100

   2.00700   -.31940   2.21500    .29520   2.27000    .26340

   2.32000   -.29840   2.39500    .00540   2.45000    .28650

   2.51900   -.04690   2.57500    .15160   2.65200    .20770

   2.70800    .10870   2.76900   -.03250   2.89300    .10330

   2.97600   -.08030   3.06800    .05200   3.12900   -.15470

   3.21200    .00650   3.25300   -.20600   3.38600    .19270

   3.41900   -.09370   3.53000    .17080   3.59900   -.03590

   3.66800    .03650   3.73800   -.07360   3.83500    .03110

   3.90400   -.18330   4.01400    .02270   4.05600   -.04350

   4.10600    .02160   4.22200   -.19720   4.31400   -.17620

   4.41600    .14600   4.47100   -.00470   4.61800    .25720

   4.66500   -.20450   4.75600    .06080   4.83100   -.27330

   4.97000    .17790   5.03900    .03010   5.10800    .21830

   5.19900    .02670   5.23300    .12520   5.30200    .12900

   5.33000    .10890   5.34300   -.02390   5.45400    .17230

   5.51000   -.10210   5.60600    .01410   5.69000   -.19490

   5.77300   -.02420   5.80000   -.00500   5.80900   -.02750

   5.86900   -.05730   5.88300   -.03270   5.92500    .02160

   5.98000    .01080   6.01300    .02350   6.08500   -.06650

   6.13200    .00140   6.17400    .04930   6.18800    .01490

   6.18810   -.02000   6.22900   -.03810   6.27900    .02070

   6.32600   -.00580   6.36800   -.06030   6.38200   -.01620

   6.40900    .02000   6.45900   -.01760   6.47800   -.00330

   6.52000    .00430   6.53400   -.00400   6.56200   -.00990

   6.57500   -.00170   6.60300   -.01700   6.64500    .03730

   6.68600    .04570   6.71400    .03850   6.72800    .00090

   6.76900   -.02880   6.76910    .00160   6.81100    .01130

   6.85200    .00220   6.90800    .00920   6.99100   -.09960

   7.07400    .03600   7.12100    .00780   7.14300   -.02770

   7.14900    .00260   7.17100    .02720   7.22600    .05760

   7.29500   -.04920   7.37000    .02970   7.40600    .01090

   7.42500    .01860   7.46100   -.02530   7.52500   -.03470

   7.57200    .00360   7.60000   -.06280   7.64100   -.02800

   7.66900   -.01960   7.69100    .00680   7.75200   -.00540

   7.79400   -.06030   7.83500   -.03570   7.87700   -.07160

   7.96000   -.01400   7.98700   -.00560   8.00100    .02220

   8.07000    .04680   8.12600    .02600   8.12610   -.03350

   8.19500   -.01280   8.22300    .06610   8.27800    .03050

   8.33400    .02460   8.40300    .03470   8.45800   -.03690

   8.53300   -.03440   8.59600   -.01040   8.63800   -.02600

   8.73500    .15340   8.81800   -.00280   8.86000    .02330

   8.88200   -.02610   8.91500   -.00220   8.95600   -.18490

   9.05300    .12600   9.09500    .03200   9.12300    .09550

   9.15000    .12460   9.25300   -.03280   9.28900   -.04510

   9.42700    .13010   9.44100   -.16570   9.51000    .04190

   9.63500   -.09360   9.70400    .08160   9.81500   -.08810

   9.89800    .00640   9.93900   -.00060   9.99500    .05860

  10.02200   -.07130  10.05000   -.04480  10.05010   -.02210

  10.10500    .00930  10.10510    .00240  10.18800    .05100

  10.27200   -.12430  10.38200    .05870  10.42400    .01330

  10.45200    .03860  10.46500    .11640  10.50700   -.03740

  10.53400   -.05720  10.64500    .03080  10.70100    .02230

  10.71400    .05150  10.77000    .09030  10.83900   -.01940

  10.92200    .04710  10.92210   -.06770  10.96400   -.07940

  10.99100   -.01200  11.07400    .06080  11.08800   -.02690

  11.11600   -.04160  11.20700    .02930  11.20710    .05520

  11.22700    .07560  11.26800    .04310  11.32400    .02080

  11.43400    .11800  11.57300   -.09990  11.65600   -.12470

  11.72500   -.20940  11.72510   -.14180  11.78000   -.11630

  11.80800   0.00000  11.87700    .07620  11.91900    .05700

  11.98800    .13540  12.04300    .06730  12.11300    .08650
```



## Release Notes

Initial release in version 14.12.

This function supersedes [SetFromFile](../../../Obsolete_Functions/SetFromFile_%7BTime_History%7D.md).

## See Also

[GetFromFile\_1](GetFromFile_1_%7BTime_History%7D.md)
