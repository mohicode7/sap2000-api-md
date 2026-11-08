# Visual Basic Concepts Used in the CSI API

Some of the Visual Basic concepts and definitions that apply to the CSI API are explained herein.

Option Base

Visual Basic 6 allows the default lower bound for arrays to be specified as 0 (the default), or 1. SAP2000 uses a lower bound of 0 for all arrays. Any program that accesses SAP2000 through the API should also use a lower bound of 0 for its arrays.

Fixed-Size and Dynamic Arrays

Arrays can be used to refer to a series of variables by the same name and to use a number (an index) to distinguish them. Visual Basic has two types of arrays: fixed-size and dynamic. A fixed-size array always remains the same size. A dynamic array can change its size while the program is running.

A fixed-size array is declared with the size indicated. For example, the following line declares MyFixedArray dimensioned to 2.



```
Dim MyFixedArray(2)as Double

 
```



Dimensioning the array to 2 means that it holds three data items:

MyFixedArray(0) = first data item

MyFixedArray(1) = second data item

MyFixedArray(2) = third data item

Dynamic arrays are declared with no size indicated as shown here:



```
Dim MyDynamicArray()as Double

 
```



Dynamic arrays are dimensioned sometime after they are declared using a statement such as the following:



```
ReDim MyDynamicArray(2)

 
```



Any array that is dimensioned inside SAP2000 must be declared as a dynamic array so that SAP2000 can redimension it. It is probably a good idea to declare all arrays as dynamic arrays for simplicity. As an example, the analysis results obtained through the CSI API are stored in arrays that are defined as dynamic arrays by the user and then dimensioned and filled inside of SAP2000.

Variable Types

Most of the data in the CSI API is one of the following variable types.

- Boolean: A variable stored as a 16-bit (2-byte) number, but it can only be True or False. When boolean values are converted to other data types, False becomes 0 and True becomes –1.
- Long: A variable stored as a 32-bit (4-byte) number ranging in value from -2,147,483,648 to 2,147,483,647. Note that other programming languages may refer to this data type differently; for example, they may refer to this as an Integer.
- Double: A double-precision floating-point variable stored as an IEEE 64-bit (8-byte) floating-point number ranging in value from -1.79769313486231E308 to -4.94065645841247E-324 for negative values and from 4.94065645841247E-324 to 1.79769313486232E308 for positive values.
- String: A variable length string.

Optional Arguments

Some of the CSI API functions have optional arguments. For example, the CountLoadDispl function has two optional arguments: Name and LoadPat. It is not necessary to include the optional arguments when calling this function. All four of the following calls are valid.



```
ret = SapModel.PointObj.CountLoadDispl(Count)



ret = SapModel.PointObj.CountLoadDispl(Count, Name)



ret = SapModel.PointObj.CountLoadDispl(Count, , LoadPat)



ret = SapModel.PointObj.CountLoadDispl(Count, Name, LoadPat)

 
```



Note in the third example, the first optional item is not included and the second optional item is included. In that case, commas must be included to denote the missing arguments.

Comments

In Visual Basic the Rem statement followed by a space indicates that all of the data on the line to the right of the Rem statement is a comment (or a remark). The Rem statement can be abbreviated using an apostrophe, '. The apostrophe is used in all of the VBA examples in the CSI API documentation to denote a comment.

ByVal and ByRef

Variables are passed to the \ using the ByRef or the ByVal keyword.

- ByVal means that the variable is passed by value. This allows the CSI API to access a copy of the variable but not the original variable. This means the value of the variable in another application can not be changed by the API.
- ByRef, which is the default in VB6 and VBA, means the argument is passed by reference. This passes the address of the variable to the CSI API instead of passing a copy of the value. It allows the CSI API to access the actual variable, and, as a result, allows SAP2000 to change the variable's actual value in an application.

Variables are passed ByRef when data needs to be returned in them from SAP2000 to your application. In addition, Visual Basic requires that all arrays be passed ByRef.

Release Notes

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

See Also

[Introduction](../Introduction.md)

[Accessing SAP2000 From An External Application](Accessing_Sap2000_From_An_External_Application.md)

[Function Documentation Conventions](Function_Documentation_Conventions.md)

[Function Return Values](Function_Return_Values.md)

[Units Abbreviations](Units_Abbreviations.md)
