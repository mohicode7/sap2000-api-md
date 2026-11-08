# Migrating From SAP2000v16 to SAP2000v17

In version 17 of SAP2000, the API has been separated from the main SAP2000 executable into a dynamic link library (DLL). This will change how API client applications connect to CSI software. You will no longer directly reference the SAP2000 executable assembly. Instead, you will reference the API DLL. In addition, you will no longer be able to declare a variable of type SapObject. Instead, create a variable of type SAP2000v17 cOAPI, which is an interface type. Then instantiate an object that implements the cOAPI interface. This process is detailed in the included examples.

# Migrating from CSiBridge16 to CSiBridge17

The same considerations and examples apply to CSiBridge 2015 (v17) as well as to SAP2000v17. In the discussion throughout this document, replace all references to SAP2000 with CSiBridge, all references to Sap2000v16 with CSiBridge16, and all references to Sap2000v17 with CSiBridge17.

See Also

Changes in API Behavior

Instructions for Updating COM clients (VBA example)

Instructions for Updating .NET clients (C# example)

[Breaking Changes](../SAP2000_API_Fuctions/Breaking_Changes_Between_v16_and_v17/Breaking_Changes.md)

[Breaking Changes to COM Enumerations](../SAP2000_API_Fuctions/Breaking_Changes_Between_v16_and_v17/Breaking_Changes_to_COM_Enumerations.md)

Breaking Changes to NET Enumerations
