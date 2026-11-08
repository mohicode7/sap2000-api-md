# cPluginContract Interface

The cPluginContract interface is used to ensure that
the cPlugin class you create contains the correct methods, so that your plugin can
be successfully added to and run by SAP2000

## Methods

### Main



```
Sub Main ( ByRef SapModel 
 As cSapModel, ByRef ISapPlugin As cPluginCallback )Dim instance As cPluginContract
 
```



This method is the entry
point for launching the plugin.

SapModel

The parameter SapModel is
a reference to the SapModel object upon which all operations will be performed.
The SapModel object will be provided to your plugin by the main SAP2000
program

ISapPlugin

The parameter ISapPlugin
is a reference to the callback object, used to retrieve the plugin status
and notify the program when the plugin is closed. The callback object
will be passed from the main SAP200 program to your plugin.

### Info



```
Function Info ( ByRef Text 
 As String ) As Integer
```



This method provides information
about the plugin

Text

The parameter Text
should be filled with information about the plugin, e.g. author, version,
description

Return Value

This method returns a zero
if the information is successfully retrieved; otherwise it returns a non-zero
value.

## See Also

[cPluginCallback
Interface](../cPlugincallback_Properties/cPluginCallback_Interfacey.md)

[Plugin
Example](../Plugin_Example_for_SAP2000_API_CHM.md)
