# cPluginCallback Interface

The cPluginCallback interface is used to check the plugin
state, and tell SAP2000 when the plugin has finished. The callback object
that implements the interface will be passed from the main SAP200 program
to your plugin.

## Methods

### Finish



```
Sub 
 Finish ( iVal As Integer)
 
```



This method must be called
when the plugin has finished. Otherwise the main SAP2000 program will
wait indefinitely.

iVal

The parameter *iVal*
sets the *ErrorFlag* property

## Properties

### ErrorFlag



```
ReadOnly
Property ErrorFlag
As Integer
 
```



This
proper indicates whether an error has occurred.

It
returns a zero if the plugin completed successfully; a non-zero
value indicates that an error occurred






### Finished



```
ReadOnly
Property Finished
As Boolean
 
```



This
property indicates whether the plugin is finished, i.e. whether the Finish method has been called.

It returns true if the plugin is finished,
otherwise, it returns false.

## See Also

[cPluginContract
Interface](../cPluginContract_Methods/cPluginContract.md)

[Plugin
Example](../Plugin_Example_for_SAP2000_API_CHM.md)
