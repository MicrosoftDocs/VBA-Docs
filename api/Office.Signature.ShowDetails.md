---
title: Signature.ShowDetails method (Office)
keywords: vbaof11.chm248014
f1_keywords:
- vbaof11.chm248014
ms.prod: office
api_name:
- Office.Signature.ShowDetails
ms.assetid: 278b84b3-c500-6357-310b-537355ad20fd
ms.date: 01/24/2019
localization_priority: Normal
---


# Signature.ShowDetails method (Office)

Displays details related to a signature packet.


## Syntax

_expression_.**ShowDetails**

_expression_ An expression that returns a **[Signature](Office.Signature.md)** object.


## Example

The following example calls the **ShowDetails** method to show details of the **Signature** object.


```vb
Sub getSignatureDetails(ByVal objSignature As Signature) 
If objSignature.IsSigned then 
 Msgbox(The document has been signed with the following details: " & objSignature.ShowDetails) 
Else 
 Msgbox("The document has not been signed.") 
End If 
End Sub 
```


## See also

- [Signature object members](overview/Library-Reference/signature-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]