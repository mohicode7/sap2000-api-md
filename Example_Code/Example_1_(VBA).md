# Example 1 (VBA)

## Remarks

This example is written for Visual Basic for Applications (VBA). It can be run from a program such as Microsoft Excel. It is based on the SAP2000 verification problem Example 1-001.

This example creates the example verification problem from scratch, runs the analysis, extracts results, and compares the results with hand-calculated values.

## Example

1. Create a new VBA project.
2. Add a reference to "SAP2000v1.tlb" to the VBA project. This early binds to the API using the COM interface.

With the release of SAP2000 v26.0.0, users will need to select the TLB file that corresponds to the bitness of their VBA development environment, e.g. 32-bit Excel or 64-bit Excel. To locate these files, the user should navigate to the directory where SAP2000 is installed. This is typically C:\Program Files\Computers and Structures\SAP2000 26\.

Within this directory, the 32-bit library is found at NativeAPI\x86\SAP2000v1.tlb The 64-bit library is found at NativeAPI\x64\SAP2000v1.tlb.

3. Add a new module to the VBA project and paste in the following code, ready-to-run code. Please pay attention to the comments in this code block; they contain important information about running the script:







```
Sub VerificationExample1001()
    Dim ret As Long

    'set the following flag to True to attach to an existing instance of the program
    'otherwise a new instance of the program will be started
    Dim AttachToInstance As Boolean
    AttachToInstance = False

    'set the following flag to True to manually specify the path to SAP2000.exe
    'this allows for a connection to a version of SAP2000 other than the latest installation
    'otherwise the latest installed version of SAP2000 will be launched
    Dim SpecifyPath As Boolean
    SpecifyPath = False

    'if the above flag is set to True, specify the path to SAP2000 below
```



    Dim ProgramPath As String



```
   ProgramPath = "C:\Program Files\Computers and Structures\SAP2000 26\SAP2000.exe"
    'full path to the model
    'set it to the desired path of your model
    Dim ModelDirectory As String
    ModelDirectory = "C:\CSiAPIexample"
    If Len(Dir(ModelDirectory, vbDirectory)) = 0 Then
        MkDir ModelDirectory
    End If

    Dim ModelName As String
    ModelName = "API_1-001.sdb"
    Dim ModelPath As String
    ModelPath = ModelDirectory & Application.PathSeparator & ModelName
   
    'create API helper object
    Dim myHelper As cHelper
    Set myHelper = New Helper

    'dimension the SapObject as cOAPI type
    Dim mySapObject As cOAPI
    Set mySapObject = Nothing
   
    If AttachToInstance Then
      'attach to a running instance of SAP2000
      'get the active SapObject
      Set mySapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject")
    Else
      If SpecifyPath Then
        'create an instance of the SapObject from the specified path
        Set mySapObject = myHelper.CreateObject(ProgramPath)
      Else
        'create an instance of the SapObject from the latest installed SAP2000
        Set mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
      End If
      'start SAP2000 application
      mySapObject.ApplicationStart
    End If
      
    'Get a reference to cSapModel to access all API classes and functions
    Dim mySapModel As cSapModel
    Set mySapModel = mySapObject.SapModel
  
    'initialize model
    ret = mySapModel.InitializeNewModel
  
    'create new blank model
    ret = mySapModel.File.NewBlank
  
    'define material property
    ret = mySapModel.PropMaterial.SetMaterial("CONC", eMatType_Concrete)
  
    'assign isotropic mechanical properties to material
    ret = mySapModel.PropMaterial.SetMPIsotropic("CONC", 3600, 0.2, 0.0000055)
  
    'define rectangular frame section property
    ret = mySapModel.PropFrame.SetRectangle("R1", "CONC", 12, 12)
  
    'define frame section property modifiers
    Dim i As Long
    Dim ModValue() As Double
    ReDim ModValue(7)
    For i = 0 To 7
        ModValue(i) = 1
    Next i
    ModValue(0) = 1000
    ModValue(1) = 0
    ModValue(2) = 0
    ret = mySapModel.PropFrame.SetModifiers("R1", ModValue)
  
    'switch to k-ft units
    ret = mySapModel.SetPresentUnits(eUnits_kip_ft_F)
  
    'add frame object by coordinates
    Dim FrameName(2) As String
    ret = mySapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, FrameName(0), "R1", "1")
    ret = mySapModel.FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, FrameName(1), "R1", "2")
    ret = mySapModel.FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, FrameName(2), "R1", "3")
  
    'assign point object restraint at base
    Dim PointName() As String
    ReDim PointName(1)
    Dim Restraint() As Boolean
    ReDim Restraint(5)
    For i = 0 To 3
        Restraint(i) = True
    Next i
  
    For i = 4 To 5
        Restraint(i) = False
    Next i
    ret = mySapModel.FrameObj.GetPoints(FrameName(0), PointName(0), PointName(1))
    ret = mySapModel.PointObj.SetRestraint(PointName(0), Restraint)
  
    'assign point object restraint at top
    For i = 0 To 1
        Restraint(i) = True
    Next i
    For i = 2 To 5
        Restraint(i) = False
    Next i
    ret = mySapModel.FrameObj.GetPoints(FrameName(1), PointName(0), PointName(1))
    ret = mySapModel.PointObj.SetRestraint(PointName(1), Restraint)
  
    'refresh view, update (initialize) zoom
    ret = mySapModel.View.RefreshView(0, False)
  
    'add load patterns
    ret = mySapModel.LoadPatterns.Add("1", eLoadPatternType_Other, 1)
    ret = mySapModel.LoadPatterns.Add("2", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("3", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("4", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("5", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("6", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("7", eLoadPatternType_Other)
  
    'assign loading for load pattern 2
    ret = mySapModel.FrameObj.GetPoints(FrameName(2), PointName(0), PointName(1))
    Dim PointLoadValue() As Double
    ReDim PointLoadValue(5)
    PointLoadValue(2) = -10
    ret = mySapModel.PointObj.SetLoadForce(PointName(0), "2", PointLoadValue)
    ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(2), "2", 1, 10, 0, 1, 1.8, 1.8)
  
    'assign loading for load pattern 3
    ret = mySapModel.FrameObj.GetPoints(FrameName(2), PointName(0), PointName(1))
    ReDim PointLoadValue(5)
    PointLoadValue(2) = -17.2
    PointLoadValue(4) = -54.4
    ret = mySapModel.PointObj.SetLoadForce(PointName(1), "3", PointLoadValue)
  
    'assign loading for load pattern 4
    ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1), "4", 1, 11, 0, 1, 2, 2)
  
    'assign loading for load pattern 5
    ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(0), "5", 1, 2, 0, 1, 2, 2, "Local")
    ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1), "5", 1, 2, 0, 1, -2, -2, "Local")
  
    'assign loading for load pattern 6
    ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(0), "6", 1, 2, 0, 1, 0.9984, 0.3744, "Local")
    ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1), "6", 1, 2, 0, 1, -0.3744, 0, "Local")
  
    'assign loading for load pattern 7
    ret = mySapModel.FrameObj.SetLoadPoint(FrameName(1), "7", 1, 2, 0.5, -15, "Local")
  
    'switch to k-in units
    ret = mySapModel.SetPresentUnits(eUnits_kip_in_F)
  
    'save model
    ret = mySapModel.File.Save(ModelPath)
  
    'run model (this will create the analysis model)
    ret = mySapModel.Analyze.RunAnalysis
  
    'initialize for results
    Dim SapResult(6) As Double
    ret = mySapModel.FrameObj.GetPoints(FrameName(1), PointName(0), PointName(1))
  
    'get results for load patterns 1 through 7
    Dim NumberResults As Long
    Dim Obj() As String
    Dim Elm() As String
    Dim LoadCase() As String
    Dim StepType() As String
    Dim StepNum() As Double
    Dim U1() As Double
    Dim U2() As Double
    Dim U3() As Double
    Dim R1() As Double
    Dim R2() As Double
    Dim R3() As Double

    For i = 0 To 6
        ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        ret = mySapModel.Results.Setup.SetCaseSelectedForOutput(Format(i + 1))
        If i <= 3 Then
            ret = mySapModel.Results.JointDispl(PointName(1), eItemTypeElm_ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
            SapResult(i) = U3(0)
        Else
            ret = mySapModel.Results.JointDispl(PointName(0), eItemTypeElm_ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
            SapResult(i) = U1(0)
        End If
    Next i
  
    'close application
    mySapObject.ApplicationExit False
    Set mySapModel = Nothing
    Set mySapObject = Nothing
  
    'fill result strings
    Dim SapResultString(6) As String
    For i = 0 To 6
        SapResultString(i) = Format(SapResult(i), "0.00000")
        If Left(SapResultString(i), 1) <> "-" Then
            SapResultString(i) = " " & SapResultString(i)
        End If
    Next i
  
    'fill independent results (hand calculated)
    Dim IndResult(6) As Double
    Dim IndResultString(6) As String
    IndResult(0) = -0.02639
    IndResult(1) = 0.06296
    IndResult(2) = 0.06296
    IndResult(3) = -0.2963
    IndResult(4) = 0.3125
    IndResult(5) = 0.11556
    IndResult(6) = 0.00651
    For i = 0 To 6
        IndResultString(i) = Format(IndResult(i), "0.00000")
        If Left(IndResultString(i), 1) <> "-" Then
            IndResultString(i) = " " & IndResultString(i)
        End If
    Next i
  
    'fill percent difference
    Dim PercentDiff(6) As Double
    Dim PercentDiffString(6) As String
    For i = 0 To 6
        PercentDiff(i) = (SapResult(i) / IndResult(i)) - 1
        PercentDiffString(i) = Format(PercentDiff(i), "0%")
        If Left(PercentDiffString(i), 1) <> "-" Then
            PercentDiffString(i) = " " & PercentDiffString(i)
        End If
    Next i
  
    'display message box comparing results
    Dim msg As String
    msg = ""
    msg = msg & "LC  Sap2000  Independent  %Diff" & vbCr & vbLf
    For i = 0 To 5
        msg = msg & Format(i + 1) & " " & SapResultString(i) & " " & IndResultString(i) & " " & PercentDiffString(i) & vbCr & vbLf
    Next i
    msg = msg & Format(i + 1) & " " & SapResultString(i) & " " & IndResultString(i) & " " & PercentDiffString(i)
    MsgBox msg

End Sub
 
```



## Release Notes

Initial release in version 15.0.1.

Updated for version 22.1.0.

Updated for version 26.0.0.
