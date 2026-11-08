# Accessing SAP2000 From An External Application

This page contains an outline for connecting to the SAP2000 API, with VBA code examples. For specific instructions for supported programming languages, please refer to the Example Code section.

The first step in using the CSI API from an external application is to reference SAP2000v1.dll or SAP2000v1.tlb from your application. If using Excel VBA, reference SAP2000v1.TLB by opening the VBA editor, clicking the Tools menu > References command and selecting SAP2000v1.TLB from the program installation folder.

Next, within your application, you will create a variable of interface type cOAPI, and an instance of the SAP2000 object which implements cOAPI. In VBA this could be accomplished as:



```
Dim mySapObject As SAP2000v1.cOAPI
Dim myHelper As SAP2000v1.cHelper
Set myHelper = New SAP2000v1.Helper

Set mySapObject = myHelper.CreateObject(ProgramPath)

 
```



The first line creates the interface variable, the second and third lines create a helper class, and the fourth line creates the instance of the SAP2000 object which implements the interface by passing in the path to where the SAP2000.exe program is located. Now that an instance of the SAP2000 object has been created in your application, start SAP2000 using the following VBA command:



```
SapObject.ApplicationStart

 
```



At this point you can open an existing model, or create a new one and perform whatever actions are required. In general, the API commands are accessed through SapObject.SapModel.

It may be helpful to define a SapModel variable so that the API commands are accessed through SapModel instead of SapObject.SapModel. In VBA this could be accomplished as:



```
Dim mySapModel As cSapModel

Set mySapModel= mySapObject.SapModel

 
```



When finished with a model, you may want to close the SAP2000 application. This can be accomplished using the following VBA command:



```
SapObject.ApplicationExit  True

 
```



As a last step, the SapModel and SapObject variables should always be set to Nothing. In VBA this is accomplished as:



```
Set SapModel= Nothing

Set SapObject= Nothing

 
```



Setting the variables to Nothing is a very important step. It breaks the connection between your application and SAP2000 and frees up system resources. If the variables are not set to Nothing, the SAP2000 application may not completely close (you may still see it running in your Windows Task Manager).

Putting all the steps previously described into a single example, a VBA program might consist of the following:



```
Sub MyProgram

   'dimension variables

      Dim mySapObject As SAP2000v1.cOAPI
      Dim myHelper As SAP2000v1.cHelper

      Dim mySapModel As cSapModel

      Dim ret As Long



   'create an instance of the SAP2000 object
      Set myHelper= New SAP2000v1.Helper

      Set mySapObject= myHelper.CreateObject("C:\Program Files (x86)\Computers and Structures\SAP2000 21\sap2000.exe")



   'start the SAP2000 application

      mySapObject.ApplicationStart



   'create the SapModel object

      Set mySapModel= mySapObject.SapModel



   'initialize model

      ret = mySapModel.InitializeNewModel



   'call SAP2000 API functions here to perform desired tasks

   'in this example a new 2D frame is created from template

      ret = mySapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'close the SAP2000 application, if desired

      mySapObject.ApplicationExit False



   'set the objects to Nothing

   'at the end of your program ALWAYS terminate the objects in this manner

      Set mySapModel= Nothing

      Set mySapObject= Nothing

End Sub
 
```



See Also

[Introduction](../Introduction.md)

[Function Documentation Conventions](Function_Documentation_Conventions.md)

[Function Return Values](Function_Return_Values.md)

[Units Abbreviations](Units_Abbreviations.md)

[Visual Basic Concepts Used in the CSI API](Visual_Basic_Concepts_Used_In_The_Sap2000_API.md)
