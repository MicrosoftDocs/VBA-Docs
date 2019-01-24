---
title: SignatureProvider.ShowSignatureDetails method (Office)
keywords: vbaof11.chm287007
f1_keywords:
- vbaof11.chm287007
ms.prod: office
api_name:
- Office.SignatureProvider.ShowSignatureDetails
ms.assetid: ea334547-af85-6d80-75dc-ddee3ad3a2c7
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureProvider.ShowSignatureDetails method (Office)

Provides a signature provider add-in the opportunity to display details about a signed signature line and display additional stored information such as a secure time-stamp.


## Syntax

_expression_.**ShowSignatureDetails**(_ParentWindow_, _psigsetup_, _psiginfo_, _XmlDsigStream_, _pcontverres_, _pcertverres_)

_expression_ An expression that returns a **[SignatureProvider](Office.SignatureProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains the handle to the window containing the signature details.|
| _psigsetup_|Required|**SignatureSetup**|Specifies initial settings of the signature provider.|
| _psiginfo_|Required|**SignatureInfo**|Specifies information about the signed signature line.|
| _XmlDsigStream_|Required|**IStream**|Represents a stream of data or binary large object of XML.|
| _pcontverres_|Required|**ContentVerificationResults**|Contains a value representing the results of verifying the signature content.|
| _pcertverres_|Required|**CertificateVerificationResults**|Contains a value representing the results of verifying the signing certification.|

## Example

The following example, written in C#, shows the implementation of the **ShowSignatureDetails** method in a custom signature provider project.


```cs
 public void ShowSignatureDetails(object parentWindow, SignatureSetup sigsetup, SignatureInfo siginfo, object xmldsigStream, ref ContentVerificationResults contverresults, ref CertificateVerificationResults certverresults) 
 { 
 using (Win32WindowFromOleWindow window = new Win32WindowFromOleWindow(parentWindow)) 
 { 
 using (SigningCeremonyForm signForm = new SigningCeremonyForm(sigsetup, siginfo)) 
 { 
 signForm.ShowDialog(window); 
 } 
 } 
 } 
 
```

> [!NOTE] 
> Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also

- [SignatureProvider object members](overview/Library-Reference/signatureprovider-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]