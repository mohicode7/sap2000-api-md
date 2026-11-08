# AddMaterial

## Syntax

SapObject.SapModel.PropMaterial.AddMaterial

## VB6 Procedure

Function AddMaterial(ByRef Name As String, ByVal MatType
As eMatType, ByVal Region As String, ByVal Standard As String, ByVal Grade
As String, Optional ByVal UserName As String = "") As Long

## Parameters

Name

This item is returned by the program. It is the name
that the program ultimately assigns for the material property. If no UserName
is specified, the program assigns a default name to the material property.
If a UserName is specified and that name is not used for another material
property, the UserName is assigned to the material property.

MatType

This is one of the following items in the eMatType enumeration.

eMatType\_Steel
= 1

eMatType\_Concrete
= 2

eMatType\_NoDesign
= 3

eMatType\_Aluminum
= 4

eMatType\_ColdFormed
= 5

eMatType\_Rebar
= 6

eMatType\_Tendon
= 7

Region

The region name of the material property that is user-predefined
in the file "CSiMaterialLibrary\*.xml" located in subfolder "Property
Libraries" under the CSiBridge installation.

Standard

The Standard name of the material property with the
specified MatType within the specified region.

Grade

The Grade name of the material property with the specified
MatType within the specified region and Standard.

UserName

This is an optional user specified name for the material
property. If a UserName is specified and that name is already used for
another material property, the program ignores the UserName.

## Remarks

This function adds a new material property to the model
based on the Code-specified and other pre-defined material properties
defined in the installed file "CSiMaterialLibrary\*.xml" located
in subfolder "Property Libraries" under the CSiBridge installation
folder.

The function returns zero if the material property is
successfully added, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddMaterial()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'add ASTM A706 rebar material property in United states 
 Region

      ret = SapModel.PropMaterial.AddMaterial(Name, 
 eMatType_Rebar, "United States", "ASTM A706", "Grade 
 60")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Updated the documentation of the eMatType enumeration
in v22.1.0

Initial release in version 15.2.0.

The function supersedes [AddQuick](../../../Obsolete_Functions/AddQuick_%7BMaterial%7D.md)
as of version 15.2.0.
