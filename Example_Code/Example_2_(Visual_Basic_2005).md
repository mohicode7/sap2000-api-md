# Example 2 (Visual Basic)

## Remarks

This Visual Basic example was created using Visual Studio 2022. It is based on the SAP2000 verification problem Example 1-001.

This example creates the example verification problem from scratch, runs the analysis, extracts results, and compares the results with hand-calculated values.

## Example

1. Create a new Visual Basic project, of type "Console App". Your project can target .NET 8 or .NET Framework 4.8, the example code is designed to be compatible with either.
2. Add a reference to SAP2000v1.dll to the project. This file is located in the directory where SAP2000 is installed, typically C:\Program Files\Computers and Structures\SAP2000 26\.
3. Copy the code below into the generated module. Please pay attention to the comments in this code block; they contain important information about running the script.v



```
Imports SAP2000v1
Module Module1
    Sub Main()
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
        Dim ProgramPath As String
        ProgramPath = "C:\Program Files\Computers and Structures\SAP2000 26\SAP2000.exe"


        'full path to the model
        'set it to the desired path of your model
        Dim ModelDirectory As String = "C:\CSiAPIexample"
        Try
            System.IO.Directory.CreateDirectory(ModelDirectory)
        Catch ex As Exception
            Console.WriteLine("Could not create directory: " + ModelDirectory)
        End Try
        Dim ModelName As String = "API_1-001.sdb"
        Dim ModelPath As String = ModelDirectory + System.IO.Path.DirectorySeparatorChar + ModelName


        'dimension the SapObject as cOAPI type
        Dim mySapObject As cOAPI
        mySapObject = Nothing


        'Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
        Dim ret As Integer
        ret = -1


        'create API helper object
        Dim myHelper As cHelper
        Try
            myHelper = New Helper
        Catch ex As Exception
            Console.WriteLine("Cannot create an instance of the Helper object")
        End Try


        If AttachToInstance Then
            'attach to a running instance of SAP2000
            Try
                'get the active SapObject
                mySapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject")
            Catch ex As Exception
                Console.WriteLine("No running instance of the program found or failed to attach.")
                Return
            End Try
        Else
            If SpecifyPath Then
                Try
                    'create an instance of the SapObject from the specified path
                    mySapObject = myHelper.CreateObject(ProgramPath)
                Catch ex As Exception
                    Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath)
                    Return
                End Try
            Else
                Try
                    'create an instance of the SapObject from the latest installed SAP2000
                    mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
                Catch ex As Exception
                    Console.WriteLine("Cannot start a new instance of the program.")
                    Return
                End Try
            End If

            'start SAP2000 application
            ret = mySapObject.ApplicationStart()
        End If


        'Get a reference to cSapModel to access all API classes and functions
        Dim mySapModel As cSapModel
        mySapModel = mySapObject.SapModel


        'initialize model
        ret = mySapModel.InitializeNewModel()


        'create new blank model
        ret = mySapModel.File.NewBlank()


        'define material property
        ret = mySapModel.PropMaterial.SetMaterial("CONC", eMatType.Concrete)


        'assign isotropic mechanical properties to material
        ret = mySapModel.PropMaterial.SetMPIsotropic("CONC", 3600, 0.2, 0.0000055)


        'define rectangular frame section property
        ret = mySapModel.PropFrame.SetRectangle("R1", "CONC", 12, 12)


        'define frame section property modifiers
        Dim ModValue(7) As Double
        For i = 0 To 7
            ModValue(i) = 1
        Next i
        ModValue(0) = 1000
        ModValue(1) = 0
        ModValue(2) = 0
        ret = mySapModel.PropFrame.SetModifiers("R1", ModValue)


        'switch to k-ft units
        ret = mySapModel.SetPresentUnits(eUnits.kip_ft_F)


        'add frame object by coordinates
        Dim FrameName(2) As String
        ret = mySapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, FrameName(0), "R1", "1")
        ret = mySapModel.FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, FrameName(1), "R1", "2")
        ret = mySapModel.FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, FrameName(2), "R1", "3")


        'assign point object restraint at base
        Dim PointName(1) As String
        Dim Restraint(5) As Boolean
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
        ret = mySapModel.LoadPatterns.Add("1", eLoadPatternType.Other, 1)
        ret = mySapModel.LoadPatterns.Add("2", eLoadPatternType.Other)
        ret = mySapModel.LoadPatterns.Add("3", eLoadPatternType.Other)
        ret = mySapModel.LoadPatterns.Add("4", eLoadPatternType.Other)
        ret = mySapModel.LoadPatterns.Add("5", eLoadPatternType.Other)
        ret = mySapModel.LoadPatterns.Add("6", eLoadPatternType.Other)
        ret = mySapModel.LoadPatterns.Add("7", eLoadPatternType.Other)


        'assign loading for load pattern 2
        ret = mySapModel.FrameObj.GetPoints(FrameName(2), PointName(0), PointName(1))
        Dim PointLoadValue(5) As Double
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
        ret = mySapModel.SetPresentUnits(eUnits.kip_in_F)


        'save model
        ret = mySapModel.File.Save(ModelPath)


        'run model (this will create the analysis model)
        ret = mySapModel.Analyze.RunAnalysis


        'initialize for results
        Dim SapResult(6) As Double
        ret = mySapModel.FrameObj.GetPoints(FrameName(1), PointName(0), PointName(1))


        'get results for load patterns 1 through 7
        Dim NumberResults As Integer = 0
        Dim Obj(0) As String
        Dim Elm(0) As String
        Dim LoadCase(0) As String
        Dim StepType(0) As String
        Dim StepNum(0) As Double
        Dim U1(0) As Double
        Dim U2(0) As Double
        Dim U3(0) As Double
        Dim R1(0) As Double
        Dim R2(0) As Double
        Dim R3(0) As Double
        For i = 0 To 6
            ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
            ret = mySapModel.Results.Setup.SetCaseSelectedForOutput(Format(i + 1))
            If i <= 3 Then
                ret = mySapModel.Results.JointDispl(PointName(1), eItemTypeElm.ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
                SapResult(i) = U3(0)
            Else
                ret = mySapModel.Results.JointDispl(PointName(0), eItemTypeElm.ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
                SapResult(i) = U1(0)
            End If
        Next i


        'close the application
        mySapObject.ApplicationExit(False)
        mySapModel = Nothing
        mySapObject = Nothing


        'fill result strings
        Dim SapResultString(6) As String
        For i = 0 To 6
            SapResultString(i) = Format(SapResult(i), "0.00000")
            If Microsoft.VisualBasic.Left(SapResultString(i), 1) <> "-" Then
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
            If Microsoft.VisualBasic.Left(IndResultString(i), 1) <> "-" Then
                IndResultString(i) = " " & IndResultString(i)
            End If
        Next i


        'fill percent difference
        Dim PercentDiff(6) As Double
        Dim PercentDiffString(6) As String
        For i As Integer = 0 To 6
            PercentDiff(i) = (SapResult(i) / IndResult(i)) - 1
            PercentDiffString(i) = Format(PercentDiff(i), "0%")
            If Microsoft.VisualBasic.Left(PercentDiffString(i), 1) <> "-" Then
                PercentDiffString(i) = " " & PercentDiffString(i)
            End If
        Next i


        'display message box comparing results
        Dim msg As String = ""
        msg = msg & "LC  SAP2000  Independent  %Diff" & vbCr & vbLf
        For i As Integer = 0 To 6
            msg = msg & Format(i + 1) & "    " & SapResultString(i) & "   " & IndResultString(i) & "       " & PercentDiffString(i) & If(i < 6, vbCr & vbLf, "")
        Next i
        Console.WriteLine(msg)


    End Sub

End Module
```



## Release Notes

Initial release in version 15.0.1.

Updated for version 22.1.0.

Updated for vesion 26.0.0.
