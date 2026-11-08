# Breaking Changes in v23.0.0, v22.2.0, and v22.1.0

The following inadvertent changes broke compatibility with earlier API versions:

In SAP2000 v22.2.0 (SAP2000v1.dll version 1.9 and CSiAPIv1.dll version 1.9), eMatTypeSteel\_Chinese\_Q345 enumeration got renamed to eMatTypeSteel\_Chinese\_Q355.

## Affected Programs:

- SAP2000 v22.2.0, v22.2.1, v23.0.0
- CSiBridge v22.2.0, v22.2.1, v23.0.0

## Affected API clients:

- Compiled COM clients (e.g. VB6, Delphi) failed to start.
- Interpreted COM clients (e.g. VBA) failed to compile/run if the affected enumeration was used.

## Fix

- eMatTypeSteel\_Chinese\_Q345 enumeration got reinstated.

In SAP2000 v23.0.0 (SAP2000v1.dll version 1.13 and CSiAPIv1.dll version 1.13), cDAlAA\_ASD\_2000 & cDAlAA\_LRFD\_2000 interfaces got removed.

## Affected Programs:

- SAP2000 v23.0.0

## Affected API clients:

- Compiled COM clients (e.g. VB6, Delphi) started but gave an automation error if affected interfaces were used.
- Interpreted COM clients (e.g. VBA) failed to compile/run if the affected interfaces were used.
- Compiled .NET clients started but gave an error if affected interfaces were used.

## Fix

- cDAlAA\_ASD\_2000 & cDAlAA\_LRFD\_2000 interfaces got reinstated.

SAP2000 versions 22.2.0, 22.2.1, and 23.0.0 and corresponding SAP2000v1.dll versions 1.9 and 1.13 & CSiAPIv1.dll versions between 1.9 and 1.13 should not to be used for developing plug-ins and/or API scripts to ensure full compatibility with past and future API versions.

cOAPI.GetOAPIVersionNumber() method can be used to check for incompatible API versions before using affected interfaces to prevent run time errors.

CSiBridge versions 22.2.0, 22.2.1, and 23.0.0 and corresponding CSiBridge1.dll versions 1.9 and 1.13 & CSiAPIv1.dll versions between 1.9 and 1.13 should not to be used for developing plug-ins and/or API scripts to ensure full compatibility with past and future API versions.
