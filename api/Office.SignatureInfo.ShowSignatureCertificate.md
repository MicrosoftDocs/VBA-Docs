---
title: SignatureInfo.ShowSignatureCertificate method (Office)
keywords: vbaof11.chm286014
f1_keywords:
- vbaof11.chm286014
ms.prod: office
api_name:
- Office.SignatureInfo.ShowSignatureCertificate
ms.assetid: 8fef7299-e110-b0a2-7a0c-552e9068e001
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureInfo.ShowSignatureCertificate method (Office)

Displays the selected or default digital certificate. 


## Syntax

_expression_.**ShowSignatureCertificate**(_ParentWindow_)

_expression_ An expression that returns a **[SignatureInfo](Office.SignatureInfo.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains the handle to the window that contains the **Certificate** dialog box.|

## Example

The following example displays a digital certificate in the window specified by the _Hwnd_ argument.

```vb
Sub DisplayCertificate(ByVal intHwnd As Long) 
Dim objSignatureInfo As SignatureInfo 
Dim objDialog As Object 
 
objDialog = objSignatureInfo.ShowSignatureCertificate(intHwnd) 
 
End Sub
```


## See also

- [SignatureInfo object members](overview/Library-Reference/signatureinfo-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]