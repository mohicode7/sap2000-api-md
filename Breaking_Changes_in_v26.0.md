# Breaking Changes in v26.0.0

[1] The Remote API, used to start and/or connect to a running instance of SAP2000 on a Remote Computer, has been disabled with the release of SAP2000 v26.0.0. API functions listed below will return an error code when called:

- cHelper.CreateObjectHost
- cHelper.CreateObjectHostPort
- cHelper.CreateObjectProgIDHost
- cHelper.CreateObjectProgIDHostPort
- cHelper.GetObjectHost
- cHelper.GetObjectHostPort

[2] In SAP2000 v26.0.0 and later, the API is compiled as a .NET Standard 2.0 AnyCPU dynamic link library (DLL). Due to that change, certain late binding VBA calls that may have worked in earlier versions of the API will no longer work. CSI recommends using early binding for VBA, and all other, API client applications, as documented in this help file.

[3] Using CSiAPIv1.dll versions 2.0 and 2.1 (released with ETABS/SAFE v22.0.0 and v22.1.0, respectively) to interact with any version of SAP2000/CSiBridge (including v26.0.0 and future versions), will result in an exception when accessing the following properties of the cLoadCases interface from an external API client/script:

- cLoadCases.HyperStatic As cCaseHyperStatic
- cLoadCases.ModalEigen As cCaseModalEigen
- cLoadCases.ModalRitz As cCaseModalRitz
- cLoadCases.ModHistLinear As cCaseModalHistoryLinear
- cLoadCases.ModHistNonlinear As cCaseModalHistoryNonlinear
- cLoadCases.ResponseSpectrum As cCaseResponseSpectrum
- cLoadCases.StaticLinear As cCaseStaticLinear
- cLoadCases.StaticNonlinear As cCaseStaticNonlinear
- cLoadCases.StaticNonlinearStaged As cCaseStaticNonlinearStaged
- cLoadCases.ExternalResults As cCaseExternalResults
- cLoadCases.Moving As cCaseMovingLoad
- cLoadCases.PSD As cCasePowerSpectralDensity
- cLoadCases.StaticLinearMultistep As cCaseStaticLinearMultistep
- cLoadCases.StaticNonlinearMultistep As cCaseStaticNonlinearMultistep
- cLoadCases.SteadyState As cCaseSteadyState

Similarly, using CSiAPIv1.dll from any version of SAP2000/CSiBridge (including v26.0.0 and future versions) to interact with ETABS/SAFE v22.0.0 and/or v22.1.0 using above properties of the cLoadCases interface from an external API client/script will result in an exception.

ETABS and SAFE versions v22.0.0 and v22.1.0 and corresponding ETABSv1.dll/SAFEv1.dll/CSiAPIv1.dll versions 2.0 and 2.1 should not to be used for developing external API clients/scripts to ensure full compatibility with past and future API versions.

cOAPI.GetOAPIVersionNumber() method can be used to check for incompatible API versions before using affected interfaces to prevent run time errors.
