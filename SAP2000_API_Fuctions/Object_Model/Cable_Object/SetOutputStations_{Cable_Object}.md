# SetOutputStations

## Syntax

SapObject.SapModel.CableObj.SetOutputStations

## VB6 Procedure

Function SetOutputStations(ByVal Name As String, ByVal MyType As Long, ByVal MaxSegSize As Double, ByVal MinSections As Long, Optional ByVal NoOutPutAndDesignAtElementEnds As Boolean, Optional ByVal NoOutPutAndDesignAtPointLoads As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

MyType

This is 1 or 2, indicating how the output stations are specified.

1 = maximum segment size, that is, maximum station spacing

2 = minimum number of stations

MaxSegSize

The maximum segment size, that is, the maximum station spacing. This item applies only when MyType = 1. [L]

MinSections

The minimum number of stations. This item applies only when MyType = 2.

NoOutPutAndDesignAtElementEnds

If this item is True, no additional output stations are added at the ends of line elements when the cable object is internally meshed.

NoOutPutAndDesignAtPointLoads

If this item is True, no additional output stations are added at point load locations.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the cable object specified by the Name item.

If this item is Group, the assignment is made to all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected cable objects, and the Name item is ignored.

## Remarks

This function assigns cable object output station data.

The function returns zero if the data is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCableOutputStationData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign cable output station data

      ret = SapModel.CableObj.SetOutputStations(Name, 1, 18, 0)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOutputStations](GetOutputStations_%7BCable_Object%7D.md)
