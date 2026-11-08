# AddByPoint

## Syntax

SapObject.SapModel.TendonObj.AddByPoint

## VB6 Procedure

Function AddByPoint(ByVal Point1 as String, ByVal Point2 as String, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "") As Long

## Parameters

Point1

The name of a defined point object at the I-End of the added tendon object.

Point2

The name of a defined point object at the J-End of the added tendon object.

Name

This is the name that the program ultimately assigns for the tendon object. If no UserName is specified, the program assigns a default name to the tendon object. If a UserName is specified and that name is not used for another frame, cable or tendon object, the UserName is assigned to the tendon object; otherwise a default name is assigned to the tendon object.

PropName

This is Default, None or the name of a defined tendon property.

If it is Default, the program assigns a default tendon property to the tendon object. If it is None, no tendon property is assigned to the tendon object. If it is the name of a defined tendon property, that property is assigned to the tendon object.

UserName

This is an optional user specified name for the tendon object. If a UserName is specified and that name is already used for another tendon object, the program ignores the UserName.

## Remarks

This function adds a new tendon object whose end points are specified by name.

The function returns zero if the tendon object is successfully added, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddTendonObjByPoint()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BTendon_Object%7D.md)

[SetTendonData](SetTendonData.md)
