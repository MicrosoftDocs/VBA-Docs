---
title: Document.SmartDocument property (Word)
keywords: vbawd10.chm158007758
f1_keywords:
- vbawd10.chm158007758
ms.prod: word
api_name:
- Word.Document.SmartDocument
ms.assetid: f9671c26-208e-1682-c792-661b701308a7
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SmartDocument property (Word)

Returns a **[SmartDocument](office.smartdocument.md)** object that represents the settings for a smart document solution.


## Syntax

_expression_.**SmartDocument**

_expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Remarks

For more information about smart documents, see [Use smart documents](https://support.office.com/article/use-smart-documents-7dc5fb76-4bc3-4583-b4ac-d94001f3dd51) on the Microsoft Support site.


## Example

The following example displays a dialog box that contains a list of valid XML expansion packs for the active document.

```vb
ActiveDocument.SmartDocument.PickSolution
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]