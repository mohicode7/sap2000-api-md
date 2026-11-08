# Introduction

The CSI Application Programming Interface (API) is a powerful tool that allows users to automate many of the processes required to build, analyze, and design models and to obtain customized analysis and design results. It also allows users to link SAP2000 or CSiBridge with third-party software, providing a path for two-way exchange of model information with other programs. Please note that SAP2000 or CSiBridge must be installed both on the API client development machine, and on the client end-user's machine.

Most major programming languages can be used to access SAP2000 and CSiBridge through the API. Detailed examples are provided for several popular programming languages, including Visual Basic, C#, Python, MATLAB, Fortran, and Visual Basic for Applications (VBA), which is included in programs such as Microsoft Excel.

With the release of SAP2000 and CSiBridge v26.0.0, the API is defined in the form of a .NET Standard 2.0 dynamic link library (DLL) for improved compatibility with a large variety of client applications. Please refer to the Release Notes page.

This documentation is organized into the following main categories:

- Release Notes explains recent changes in the API, and include instructions for users on how to update their client applications.
- **Getting Started** briefly explains how to use the CSI API and how the CSI API functions are documented.
- **CSI API Functions** identifies each function available in the API and provides an example of how the function might be called using Visual Basic for Applications (VBA).
- **Example Code** provides programming examples using the CSI API. These examples are more extensive than those included with the documentation of each function.
- Obsolete Functions are the result of changes to the software or the API. These obsolete functions have been superseded, but continue to be included to accommodate backwards compatibility.
- Breaking Changes lists recent changes to the API that might break some API clients.
- Plugins explains how to develop and use plugins that can be accessed via program menus.
- Database Tables (Interactive Editing) documents each function available in DatabaseTables class for accessing database tables through the API.
- Alphabetized List of Functions lists all function in alphabetical order.

See Also

[Accessing SAP2000 From An External Application](Getting_Started/Accessing_Sap2000_From_An_External_Application.md)

[Function Documentation Conventions](Getting_Started/Function_Documentation_Conventions.md)

[Function Return Values](Getting_Started/Function_Return_Values.md)

[Units Abbreviations](Getting_Started/Units_Abbreviations.md)

[Visual Basic Concepts Used In the CSI API](Getting_Started/Visual_Basic_Concepts_Used_In_The_Sap2000_API.md)
