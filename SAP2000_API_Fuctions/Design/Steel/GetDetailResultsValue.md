# **GetDetailResultsValue**

## **Syntax**

SapObject.SapModel.DesignSteel.GetDetailResultsValue

## **VB6 Procedure**

Public Function GetDetailResultsValue(ByVal Name As String, ByVal ItemType As eItemType, ByVal Table As Long, ByVal Field As String, ByRef NumberItems As Long, FrameName() As String, Text() As String) As Long

## **Parameters**

**Name**

The name of an existing frame object or group, depending on the value of the ItemType item.

**ItemType**

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the design results are retrieved for the frame object specified by the Name item.

If this item is Group, the design results are retrieved for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the design results are retrieved for all selected frame objects, and the Name item is ignored.

**Table**

The table ID of the steel design output database Tables. The table names are input as the representative table numbers and are code-based. Please see the following appendix.

**Field**

The field name with Numerical output data type in the specified steel design result database Tables. The Field Names need to be the exactly same as the names in the specified steel design output database tables except the case is insensitive.

**NumberItems**

The number of frame objects for which results are obtained.

**FrameName**

The frame object names for which results are obtained.

**Value**

The design results with Numerical output data type of the request field in the request table for the specified frame objects.

## **Remarks**

This function retrieves the design results from steel design output database tables. Note that the summary table of all design codes is not included in this function.

The function returns zero if the results are successfully retrieved; otherwise it returns a nonzero value.

## **VBA Example**

Sub GetSteelDesignDetailResultsValue ()  
'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim NumberItems As Long

Dim FrameName() As String

Dim Text() As String

'create Sap2000 object

Set SapObject = New Sap2000v16.SapObject

'start Sap2000 application

SapObject.ApplicationStart

'create SapModel object

Set SapModel = SapObject.SapModel

'initialize model

ret = SapModel.InitializeNewModel

'create model from template

ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)

'run analysis

ret = SapModel.File.Save("C:\SapAPI\x.sdb")

ret = SapModel.Analyze.RunAnalysis

'start steel design

ret = SapModel.DesignSteel.SetCode("AISC360-05/IBC2006")

ret = SapModel.DesignSteel.StartDesign

'get summary result data

ret = SapModel.DesignSteel. GetDetailResultsValue("8", 0, 2, "Pr", NumberItems, FrameName(), Value())

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## **Release Notes**

Initial release in version 16.0.0.

## **See Also**

[GetDetailResultsText](GetDetailResultsText.md)

## **Appendix – The available table index for selected design codes**

"AISC360-05/IBC2006"

  2:  "Steel Design 2 - PMM Details - AISC360-05-IBC2006"

  3:  "Steel Design 3 - Shear Details - AISC360-05-IBC2006"

  4:  "Steel Design 4 - Continuity Plates - AISC360-05-IBC2006"

  5:  "Steel Design 5 - Doubler Plates - AISC360-05-IBC2006"

  6:  "Steel Design 6 - Beam/Column Ratios - AISC360-05-IBC2006"

  7:  "Steel Design 7 - Beam Shear Forces - AISC360-05-IBC2006"

  8:  "Steel Design 8 - Brace Max Axial Load - AISC360-05-IBC2006"

  9:  "Steel Design 9 - Decision Parameters - AISC360-05-IBC2006"

"AISC-ASD89"

  2:  "Steel Design 2 - PMM Details - AISC-ASD89"

  3:  "Steel Design 3 - Shear Details - AISC-ASD89"

  4:  "Steel Design 4 - Continuity Plates - AISC-ASD89"

  5:  "Steel Design 5 - Doubler Plates - AISC-ASD89"

  6:  "Steel Design 6 - Beam/Column Ratios - AISC-ASD89"

  7:  "Steel Design 7 - Beam Shear Forces - AISC-ASD89"

  8:  "Steel Design 8 - Brace Max Axial Load - AISC-ASD89"

"AISC-ASD01"

  2:  "Steel Design 2 - PMM Details - AISC-ASD01"

  3:  "Steel Design 3 - Shear Details - AISC-ASD01"

  4:  "Steel Design 4 - Continuity Plates - AISC-ASD01"

  5:  "Steel Design 5 - Doubler Plates - AISC-ASD01"

  6:  "Steel Design 6 - Beam/Column Ratios - AISC-ASD01"

  7:  "Steel Design 7 - Beam Shear Forces - AISC-ASD01"

  8:  "Steel Design 8 - Brace Max Axial Load - AISC-ASD01"

"AISC-LRFD99"

  2:  "Steel Design 2 - PMM Details - AISC-LRFD99"

  3:  "Steel Design 3 - Shear Details - AISC-LRFD99"

  4:  "Steel Design 4 - Continuity Plates - AISC-LRFD99"

  5:  "Steel Design 5 - Doubler Plates - AISC-LRFD99"

  6:  "Steel Design 6 - Beam/Column Ratios - AISC-LRFD99"

  7:  "Steel Design 7 - Beam Shear Forces - AISC-LRFD99"

  8:  "Steel Design 8 - Brace Max Axial Load - AISC-LRFD99"

"AISC-LRFD93"

  2:  "Steel Design 2 - PMM Details - AISC-LRFD93"

  3:  "Steel Design 3 - Shear Details - AISC-LRFD93"

  4:  "Steel Design 4 - Continuity Plates - AISC-LRFD93"

  5:  "Steel Design 5 - Doubler Plates - AISC-LRFD93"

  6:  "Steel Design 6 - Beam/Column Ratios - AISC-LRFD93"

  7:  "Steel Design 7 - Beam Shear Forces - AISC-LRFD93"

  8:  "Steel Design 8 - Brace Max Axial Load - AISC-LRFD93"

"AASHTO LRFD 2007"

  2:  "Steel Design 2 - PMM Details - AASHTO-LRFD-2007"

  3:  "Steel Design 3 - Shear Details - AASHTO-LRFD-2007"

  4:  "Steel Design 4 - Continuity Plates - AASHTO-LRFD-2007"

  5:  "Steel Design 5 - Doubler Plates - AASHTO-LRFD-2007"

  6:  "Steel Design 6 - Beam/Column Ratios - AASHTO-LRFD-2007"

  7:  "Steel Design 7 - Beam Shear Forces - AASHTO-LRFD-2007"

  8:  "Steel Design 8 - Brace Max Axial Load - AASHTO-LRFD-2007"

"API RP2A-LRFD 97"

  2:  "Steel Design 2 - PMM Details - API RP2A-LRFD 97"

  3:  "Steel Design 2 - PMM Details for Pipes - API RP2A-LRFD 97"

  4:  "Steel Design 3 - Shear Details - API RP2A-LRFD 97"

  5:  "Steel Design 4 - Continuity Plates - API RP2A-LRFD 97"

  6:  "Steel Design 5 - Doubler Plates - API RP2A-LRFD 97"

  7:  "Steel Design 6 - Beam/Column Ratios - API RP2A-LRFD 97"

  8:  "Steel Design 7 - Beam Shear Forces - API RP2A-LRFD 97"

  9:  "Steel Design 8 - Brace Max Axial Load - API RP2A-LRFD 97"

"API RP2A-WSD2000"

  2:  "Steel Design 2 - PMM Details - API RP2A-WSD2000"

  3:  "Steel Design 2 - PMM Details for Pipes - API RP2A-WSD2000"

  4:  "Steel Design 3 - Shear Details - API RP2A-WSD2000"

  5:  "Steel Design 4 - Continuity Plates - API RP2A-WSD2000"

  6:  "Steel Design 5 - Doubler Plates - API RP2A-WSD2000"

  7:  "Steel Design 6 - Beam/Column Ratios - API RP2A-WSD2000"

  8:  "Steel Design 7 - Beam Shear Forces - API RP2A-WSD2000"

  9:  "Steel Design 8 - Brace Max Axial Load - API RP2A-WSD2000"

"AS 4100-1998"

  2:  "Steel Design 2 - PMM Details - AS 4100-1998"

  3:  "Steel Design 3 - Shear Details - AS 4100-1998"

  4:  "Steel Design 4 - Continuity Plates - AS 4100-1998"

  5:  "Steel Design 5 - Doubler Plates - AS 4100-1998"

  6:  "Steel Design 6 - Beam/Column Ratios - AS 4100-1998"

  7:  "Steel Design 7 - Beam Shear Forces - AS 4100-1998"

  8:  "Steel Design 8 - Brace Max Axial Load - AS 4100-1998"

"ASCE 10-97"

  2:  "Steel Design 2 - PMM Details - ASCE 10-97"

  3:  "Steel Design 2 - PMM Details for Angles - ASCE 10-97"

  4:  "Steel Design 3 - Shear Details - ASCE 10-97"

  5:  "Steel Design 3 - Shear Details for Angles - ASCE 10-97"

  6:  "Steel Design 4 - Continuity Plates - ASCE 10-97"

  7:  "Steel Design 5 - Doubler Plates - ASCE 10-97"

  8:  "Steel Design 6 - Beam/Column Ratios - ASCE 10-97"

  9:  "Steel Design 7 - Beam Shear Forces - ASCE 10-97"

  10: "Steel Design 8 - Brace Max Axial Load - ASCE 10-97"

"EUROCODE 3-1993"

  2:  "Steel Design 2 - PMM Details - EUROCODE 3-1993"

  3:  "Steel Design 3 - Shear Details - EUROCODE 3-1993"

  4:  "Steel Design 4 - Continuity Plates - EUROCODE 3-1993"

  5:  "Steel Design 5 - Doubler Plates - EUROCODE 3-1993"

  6:  "Steel Design 6 - Beam/Column Ratios - EUROCODE 3-1993"

  7:  "Steel Design 7 - Beam Shear Forces - EUROCODE 3-1993"

  8:  "Steel Design 8 - Brace Max Axial Load - EUROCODE 3-1993"

"Eurocode 3-2005"

  2:  "Steel Design 2 - PMM Details - Eurocode 3-2005"

  3:  "Steel Design 3 - Shear Details - Eurocode 3-2005"

  4:  "Steel Design 4 - Continuity Plates - Eurocode 3-2005"

  5:  "Steel Design 5 - Doubler Plates - Eurocode 3-2005"

  6:  "Steel Design 6 - Beam/Column Ratios - Eurocode 3-2005"

  7:  "Steel Design 7 - Beam Shear Forces - Eurocode 3-2005"

  8:  "Steel Design 8 - Brace Max Axial Load - Eurocode 3-2005"

"Indian IS 800:2007"

  2:  "Steel Design 2 - PMM Details - Indian IS 800:2007"

  3:  "Steel Design 3 - Shear Details - Indian IS 800:2007"

  4:  "Steel Design 4 - Continuity Plates - Indian IS 800:2007"

  5:  "Steel Design 5 - Doubler Plates - Indian IS 800:2007"

  6:  "Steel Design 6 - Beam/Column Ratios - Indian IS 800:2007"

  7:  "Steel Design 7 - Beam Shear Forces - Indian IS 800:2007"

  8:  "Steel Design 8 - Brace Max Axial Load - Indian IS 800:2007"

"Italian UNI 10011"

  2:  "Steel Design 2 - PMM Details - Italian UNI 10011"

  3:  "Steel Design 3 - Shear Details - Italian UNI 10011"

  4:  "Steel Design 4 - Continuity Plates - Italian UNI 10011"

  5:  "Steel Design 5 - Doubler Plates - Italian UNI 10011"

  6:  "Steel Design 6 - Beam/Column Ratios - Italian UNI 10011"

  7:  "Steel Design 7 - Beam Shear Forces - Italian UNI 10011"

  8:  "Steel Design 8 - Brace Max Axial Load - Italian UNI 10011"

"NZS 3404-1997"

  2:  "Steel Design 2 - PMM Details - NZS 3404-1997"

  3:  "Steel Design 3 - Shear Details - NZS 3404-1997"

  4:  "Steel Design 4 - Continuity Plates - NZS 3404-1997"

  5:  "Steel Design 5 - Doubler Plates - NZS 3404-1997"

  6:  "Steel Design 6 - Beam/Column Ratios - NZS 3404-1997"

  7:  "Steel Design 7 - Beam Shear Forces - NZS 3404-1997"

  8:  "Steel Design 8 - Brace Max Axial Load - NZS 3404-1997"

"Norsok N-004"

  2:  "Steel Design 2 - PMM Details - Norsok N-004"

  3:  "Steel Design 2 - PMM Details for Pipes - Norsok N-004"

  4:  "Steel Design 3 - Shear Details - Norsok N-004"

  5:  "Steel Design 4 - Continuity Plates - Norsok N-004"

  6:  "Steel Design 5 - Doubler Plates - Norsok N-004"

  7:  "Steel Design 6 - Beam/Column Ratios - Norsok N-004"

  8:  "Steel Design 7 - Beam Shear Forces - Norsok N-004"

  9:  "Steel Design 8 - Brace Max Axial Load - Norsok N-004"

"UBC97-ASD"

  2:  "Steel Design 2 - PMM Details - UBC97-ASD"

  3:  "Steel Design 3 - Shear Details - UBC97-ASD"

  4:  "Steel Design 4 - Continuity Plates - UBC97-ASD"

  5:  "Steel Design 5 - Doubler Plates - UBC97-ASD"

  6:  "Steel Design 6 - Beam/Column Ratios - UBC97-ASD"

  7:  "Steel Design 7 - Beam Shear Forces - UBC97-ASD"

  8:  "Steel Design 8 - Brace Max Axial Load - UBC97-ASD"

"UBC97-LRFD"

  2:  "Steel Design 2 - PMM Details - UBC97-LRFD"

  3:  "Steel Design 3 - Shear Details - UBC97-LRFD"

  4:  "Steel Design 4 - Continuity Plates - UBC97-LRFD"

  5:  "Steel Design 5 - Doubler Plates - UBC97-LRFD"

  6:  "Steel Design 6 - Beam/Column Ratios - UBC97-LRFD"

  7:  "Steel Design 7 - Beam Shear Forces - UBC97-LRFD"

  8:  "Steel Design 8 - Brace Max Axial Load - UBC97-LRFD"
