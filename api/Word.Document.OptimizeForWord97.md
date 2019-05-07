---
title: Document.OptimizeForWord97 property (Word)
keywords: vbawd10.chm158007630
f1_keywords:
- vbawd10.chm158007630
ms.prod: word
api_name:
- Word.Document.OptimizeForWord97
ms.assetid: 9db75633-508c-eddb-1ee9-5c8a2e9969b2
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.OptimizeForWord97 property (Word)

 **True** if Microsoft Word optimizes the current document for viewing in Microsoft Word 97 by disabling any incompatible formatting. Read/write **Boolean**.


## Syntax

_expression_. `OptimizeForWord97`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

To optimize all new documents for Word 97 by default, use the  **[OptimizeForWord97byDefault](Word.Options.OptimizeForWord97byDefault.md)** property.


## Example

This example checks the current document to see if it is optimized for Word 97; if it is not, the example asks the user whether it should be.


```vb
If ActiveDocument.OptimizeForWord97 = False Then 
 x = MsgBox("Is this document targeted at " _ 
 & "Word 97 users?", vbYesNo) 
 If x = vbYes Then _ 
 ActiveDocument.OptimizeForWord97 = True 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]