# Interactive Database

Beginning
with Version 23.0.0, SAP2000 and CSIBridge allow database tables to be
accessed through the API using the DatabaseTables
class.
It is recommended that API users first familiarize themselves with the
database tables using the SAP2000 (or CSiBridge) GUI, to understand how
data can be requested, viewed, modified, and applied to the model.

Generally,
the API user will start with a call
to [GetAvailableTables](../Database_Tables/GetAvailableTables.md).
They can decide which tables they want to retrieve data from, or edit
and apply to the model. A call to [GetAllFieldsInTable](../Database_Tables/GetAllFieldsInTable.md)
will
show the user which columns the table has, and which columns can be edited
and imported.

If
the user simply wants to retrieve data, they can use one of the
[GetTableForDisplay](../Database_Tables/GetTableforDisplayArray.md)...functions.
Functions are provided to set the Load Cases, Load Combinations, Load
Patterns, Generalized Displacements, Section Cuts, Pushover Named Sets,
Plot Function Traces Named Sets, Joint Response Spectra Named Sets, and
Element Virtual Work Named Sets for which data is desired. The
[SetTableOutputOptionsForDisplay](../Database_Tables/GetTableOutputOptionsForDisplay.md)
function
can be used to set other options for display.

If
the user would like to edit a table and import it into the model, they
should start by calling one of the [GetTableForEditing](../Database_Tables/GetTableForEditingArray.md)...functions.
This will retrieve the table data in one of several formats. The user
can then edit that data, but they must ensure that the format of the data
is not altered. Then they can import the edited table data with the respective
[SetTableForEditing](../Database_Tables/SetTableforEditingArray.md)...function.
These functions only operate on one table at a time, but any number of
edited tables can be imported by calling them successively. Finally, to
apply the edited tables to the model, the user will call [ApplyEditedTables](../Database_Tables/ApplyEditedTables.md).
If for any reason the user would like to clear the internal list of edited
tables set using the [SetTableForEditing](../Database_Tables/SetTableforEditingArray.md)...
functions,
they can call the [CancelTableEditing](../Database_Tables/CancelTableEditing.md)
function.

For
thorough descriptions of database table operations, please refer to the
function documentation in DatabaseTables
