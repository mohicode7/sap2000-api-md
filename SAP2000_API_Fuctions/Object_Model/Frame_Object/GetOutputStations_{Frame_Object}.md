# GetOutputStations

## Syntax

SapObject.SapModel.FrameObj.GetOutputStations

## VB6 Procedure

Function GetOutputStations(ByVal Name As String, ByRef MyType As Long, ByRef MaxSegSize As Double, ByRef MinSections As Long, ByRef NoOutPutAndDesignAtElementEnds As Boolean, ByRef NoOutPutAndDesignAtPointLoads As Boolean) As Long

## Parameters

Name

The name of an existing frame object.

MyType

This is either 1 or 2 indicating how the output stations are specified.

1 = maximum segment size, that is, maximum station spacing

2 = minimum number of stations

MaxSegSize

The maximum segment size, that is, the maximum station spacing. This item applies only when MyType = 1. [L]

MinSections

The minimum number of stations. This item applies only when MyType = 2.

NoOutPutAndDesignAtElementEnds

If this item is True, no additional output stations are added at the ends of line elements when the frame object is internally meshed.

NoOutPutAndDesignAtPointLoads

If this item is True, no additional output stations are added at point load locations.

## Remarks

This function retrieves frame object output station data.

The function returns zero if the data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameOutputStationData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim MaxSegSize As Double

      Dim MinSections As Long

      Dim NoOutPutAndDesignAtElementEnds As Boolean

      Dim NoOutPutAndDesignAtPointLoads As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'get frame output station data

      ret = SapModel.FrameObj.GetOutputStations("15", MyType, MaxSegSize, MinSections, NoOutPutAndDesignAtElementEnds, NoOutPutAndDesignAtPointLoads)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetOutputStations](SetOutputStations_%7BFrame_Object%7D.md)
