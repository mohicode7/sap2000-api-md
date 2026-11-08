# AddQuick  (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropMaterial.AddQuick

## VB6 Procedure

Function AddQuick(ByVal Name As String, ByVal MatType As eMatType, Optional ByVal SteelType As eMatTypeSteel = MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A992\_Fy50, Optional ByVal ConcreteType As eMatTypeConcrete = MATERIAL\_CONCRETE\_SUBTYPE\_FC4000\_NORMALWEIGHT, Optional ByVal AluminumType As eMatTypeAluminum = MATERIAL\_ALUMINUM\_SUBTYPE\_6061\_T6, Optional ByVal ColdFormedType As eMatTypeColdFormed = MATERIAL\_COLDFORMED\_SUBTYPE\_ASTM\_A653SQGr50, Optional ByVal RebarType As eMatTypeRebar = MATERIAL\_REBAR\_SUBTYPE\_ASTM\_A615Gr60, Optional ByVal TendonType As eMatTypeTendon = MATERIAL\_TENDON\_SUBTYPE\_ASTM\_A416Gr270, Optional ByVal UserName As String = "") As Long

## Parameters

Name

This item is returned by the program. It is the name that the program ultimately assigns for the material property. If no UserName is specified, the program assigns a default name to the material property. If a UserName is specified and that name is not used for another material property, the UserName is assigned to the material property.

MatType

This is one of the following items in the eMatType enumeration.

MATERIAL\_STEEL = 1

MATERIAL\_CONCRETE = 2

MATERIAL\_NODESIGN = 3

MATERIAL\_ALUMINUM = 4

MATERIAL\_COLDFORMED = 5

MATERIAL\_REBAR = 6

MATERIAL\_TENDON = 7

SteelType

This is one of the following items in the eMatTypeSteel enumeration.

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A36 = 1

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A53GrB = 2

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A500GrB\_Fy42 = 3

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A500GrB\_Fy46 = 4

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A572Gr50 = 5

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A913Gr50 = 6

MATERIAL\_STEEL\_SUBTYPE\_ASTM\_A992\_Fy50 = 7

MATERIAL\_STEEL\_SUBTYPE\_CHINESE\_Q235 = 8

MATERIAL\_STEEL\_SUBTYPE\_CHINESE\_Q345 = 9

MATERIAL\_STEEL\_SUBTYPE\_INDIAN\_Fe250 = 10

MATERIAL\_STEEL\_SUBTYPE\_INDIAN\_Fe345 = 11

MATERIAL\_STEEL\_SUBTYPE\_EN100252\_S235 = 12

MATERIAL\_STEEL\_SUBTYPE\_EN100252\_S275 = 13

MATERIAL\_STEEL\_SUBTYPE\_EN100252\_S355 = 14

MATERIAL\_STEEL\_SUBTYPE\_EN100252\_S450 = 15

This item is applicable only when MatType = MATERIAL\_STEEL.

ConcreteType

This is one of the following items in the eMatTypeConcrete enumeration.

MATERIAL\_CONCRETE\_SUBTYPE\_FC3000\_NORMALWEIGHT = 1

MATERIAL\_CONCRETE\_SUBTYPE\_FC4000\_NORMALWEIGHT = 2

MATERIAL\_CONCRETE\_SUBTYPE\_FC5000\_NORMALWEIGHT = 3

MATERIAL\_CONCRETE\_SUBTYPE\_FC6000\_NORMALWEIGHT = 4

MATERIAL\_CONCRETE\_SUBTYPE\_FC3000\_LIGHTWEIGHT = 5

MATERIAL\_CONCRETE\_SUBTYPE\_FC4000\_LIGHTWEIGHT = 6

MATERIAL\_CONCRETE\_SUBTYPE\_FC5000\_LIGHTWEIGHT = 7

MATERIAL\_CONCRETE\_SUBTYPE\_FC6000\_LIGHTWEIGHT = 8

MATERIAL\_CONCRETE\_SUBTYPE\_CHINESE\_C20\_NORMALWEIGHT = 9

MATERIAL\_CONCRETE\_SUBTYPE\_CHINESE\_C30\_NORMALWEIGHT = 10

MATERIAL\_CONCRETE\_SUBTYPE\_CHINESE\_C40\_NORMALWEIGHT = 11

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M15\_NORMALWEIGHT = 12

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M20\_NORMALWEIGHT = 13

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M25\_NORMALWEIGHT = 14

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M30\_NORMALWEIGHT = 15

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M35\_NORMALWEIGHT = 16

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M40\_NORMALWEIGHT = 17

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M45\_NORMALWEIGHT = 18

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M50\_NORMALWEIGHT = 19

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M55\_NORMALWEIGHT = 20

MATERIAL\_CONCRETE\_SUBTYPE\_INDIAN\_M60\_NORMALWEIGHT = 21

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C12\_NORMALWEIGHT = 22

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C16\_NORMALWEIGHT = 23

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C20\_NORMALWEIGHT = 24

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C25\_NORMALWEIGHT = 25

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C30\_NORMALWEIGHT = 26

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C35\_NORMALWEIGHT = 27

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C40\_NORMALWEIGHT = 28

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C45\_NORMALWEIGHT = 29

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C50\_NORMALWEIGHT = 30

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C55\_NORMALWEIGHT = 31

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C60\_NORMALWEIGHT = 32

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C70\_NORMALWEIGHT = 33

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C80\_NORMALWEIGHT = 34

MATERIAL\_CONCRETE\_SUBTYPE\_EN\_C90\_NORMALWEIGHT = 35

This item is applicable only when MatType = MATERIAL\_CONCRETE.

AluminumType

This is one of the following items in the eMatTypeAluminum enumeration.

MATERIAL\_ALUMINUM\_SUBTYPE\_6061\_T6 = 1

MATERIAL\_ALUMINUM\_SUBTYPE\_6063\_T6 = 2

MATERIAL\_ALUMINUM\_SUBTYPE\_5052\_H34 = 3

This item is applicable only when MatType = MATERIAL\_ALUMINUM.

ColdFormedType

This is one of the following items in the eMatTypeColdFormed enumeration.

MATERIAL\_COLDFORMED\_SUBTYPE\_ASTM\_A653SQGr33 = 1

MATERIAL\_COLDFORMED\_SUBTYPE\_ASTM\_A653SQGr50 = 2

This item is applicable only when MatType = MATERIAL\_COLDFORMED.

RebarType

This is one of the following items in the eMatTypeRebar enumeration.

MATERIAL\_REBAR\_SUBTYPE\_ASTM\_A615Gr40 = 1

MATERIAL\_REBAR\_SUBTYPE\_ASTM\_A615Gr60 = 2

MATERIAL\_REBAR\_SUBTYPE\_ASTM\_A615Gr75 = 3

MATERIAL\_REBAR\_SUBTYPE\_ASTM\_A706 = 4

MATERIAL\_REBAR\_SUBTYPE\_CHINESE\_HPB235 = 5

MATERIAL\_REBAR\_SUBTYPE\_CHINESE\_HRB335 = 6

MATERIAL\_REBAR\_SUBTYPE\_CHINESE\_HRB400 = 7

MATERIAL\_REBAR\_SUBTYPE\_INDIAN\_Mild250 = 8

MATERIAL\_REBAR\_SUBTYPE\_INDIAN\_HYSD415 = 9

MATERIAL\_REBAR\_SUBTYPE\_INDIAN\_HYSD500 = 10

MATERIAL\_REBAR\_SUBTYPE\_INDIAN\_HYSD550 = 11

This item is applicable only when MatType = MATERIAL\_REBAR.

TendonType

This is one of the following items in the eMatTypeTendon enumeration.

MATERIAL\_TENDON\_SUBTYPE\_ASTM\_A416Gr250 = 1

MATERIAL\_TENDON\_SUBTYPE\_ASTM\_A416Gr270 = 2

This item is applicable only when MatType = MATERIAL\_TENDON.

UserName

This is an optional user specified name for the material property. If a UserName is specified and that name is already used for another material property, the program ignores the UserName.

## Remarks

This function adds a new material property to the model using built-in default values.

The function returns zero if the property is successfully added; otherwise it returns nonzero.

## VBA Example



```
Sub AddMaterialQuick()

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



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Added Steel Material Types for Indian and European codes, Concrete Material Types for Indian and European codes, and Rebar for the Indian code in SAP2000 Version 15.0.0 and CSiBridge Version 15.1.0.

The function is obsolete and has been superseded by [AddMaterial](../Definitions/Properties/Material/AddMaterial.md) as of version 15.2.0. This function is maintained for backward compatibility. New function added.

## See Also
