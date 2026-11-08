# Creating an API client in .NET 6 (for programs older than v26.0.0)

The SAP2000 (or CSiBridge) API is provided as a .NET Framework 4.7.1 class library. The .NET remoting libraries used by the SAP2000 API are no longer supported in .NET 6 . The Visual Basic or C# examples provided in this help file can be compiled in a .NET 6 client but will produce a runtime error.

Users who must create a .NET 6 client can still access the SAP2000 API using COM. This can be done by changing one line in the example code:

## Original Visual Basic Example:



```
 'create API helper object
 Dim myHelper As cHelper
 Try
     myHelper = New Helper

 Catch ex As Exception

     MsgBox("Cannot create an instance of the Helper object")

 End Try
```



## Modified Visual Basic Example:



```
'create API helper object
Dim myHelper As cHelper
Try
      myHelper = Activator.CreateInstance(Type.GetTypeFromProgID("SAP2000v1.Helper", True))

Catch ex As Exception
      MsgBox("Cannot create an instance of the Helper object")

End Try
```



## Original C# Example:



```
//create API helper object
cHelper myHelper;
try
{
    myHelper = new Helper();
}
catch (Exception ex)
{
    Console.WriteLine("Cannot create an instance of the Helper object");
    return;
}
```



## Modified C# Example:



```
//create API helper object
cHelper myHelper;
try
{
    myHelper = (cHelper)Activator.CreateInstance(Type.GetTypeFromProgID("SAP2000v1.Helper", true));
}
catch (Exception ex)
{
    Console.WriteLine("Cannot create an instance of the Helper object");
    return;
}

```
