---
title: SignatureInfo.SelectSignatureCertificate method (Office)
keywords: vbaof11.chm286015
f1_keywords:
- vbaof11.chm286015
ms.prod: office
api_name:
- Office.SignatureInfo.SelectSignatureCertificate
ms.assetid: acf3993f-85b3-a455-e3ee-1a713e7787c6
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureInfo.SelectSignatureCertificate method (Office)

Displays a dialog box that allows users to select which signature certificate to use for signing a document.


## Syntax

_expression_.**SelectSignatureCertificate**(_ParentWindow_)

_expression_ An expression that returns a **[SignatureInfo](Office.SignatureInfo.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains a handle to the window containing the certificate selection dialog box.|

## Example

The following example displays a dialog box that allows the user to select a digital certificate.


```vb
Sub SelectDigCertificate(ByVal intHwnd As Long) 
Dim objSignatureInfo As SignatureInfo 
Dim objDialog As Object 
 
objDialog = objSignatureInfo.SelectSignatureCertificate(intHwnd) 
 
End Sub
```


## See also

- [SignatureInfo object members](overview/Library-Reference/signatureinfo-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]