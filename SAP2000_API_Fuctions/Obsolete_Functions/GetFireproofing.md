# GetFireproofing

## Syntax

SapObject.SapModel.FrameObj.GetFireproofing

## VB6 Procedure

Function GetFireproofing(ByVal Name As String, ByRef MyType As Long, ByRef Thickness As Double, ByRef Perimeter As Double, ByRef Density As Double, ByRef tf As Boolean) As Long

## Parameters

Name

The name of an existing frame object.

MyType

This is 1, 2 or 3, indicating the type of fireproofing assigned.

1 = Sprayed on - program calculate section perimeter

2 = Sprayed on - user provides section perimeter

3 = Concrete encased

Thickness

When MyType = 1 or MyType = 2 this is the thickness of the sprayed on fireproofing. When MyType = 3 this is the concrete cover dimension. [L]

Perimeter

This item applies only when MyType = 2. It is the length of fireproofing applied measured around the perimeter of the frame object cross-section. [L]

Density

This is the weight per unit volume of the fireproofing material. [F/L3]

tf

This item  applies only when MyType = 1 or MyType = 3. If this item is True, the fireproofing is assumed to be applied to the top flange of the section. If it is False, the program assumes no fireproofing is applied to the section top flange. This flag applies for I, channel and double channel sections.

## Remarks

This function retrieves the fireproofing assignments to frame objects.

The function returns zero if the fireproofing assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFireproofing()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim Thickness As Double

      Dim Perimeter As Double

      Dim Density As Double

      Dim tf As Boolean



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



   'assign fireproofing

      ret = SapModel.FrameObj.SetFireproofing("ALL", 1, 2, 0, 8.68E-06, False, Group)



   'get fireproofing

      ret = SapModel.FrameObj.GetFireproofing("3", MyType, Thickness, Perimeter, Density, tf)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetFireproofing](SetFireproofing.md)

[DeleteFireproofing](../Object_Model/Frame_Object/DeleteFireproofing.md)
