---
title: Document.ScratchArea property (Publisher)
keywords: vbapb10.chm196657
f1_keywords:
- vbapb10.chm196657
ms.prod: publisher
api_name:
- Publisher.Document.ScratchArea
ms.assetid: 782d9b7f-b620-60f0-c21d-04f588c37cc6
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ScratchArea property (Publisher)

Returns a **[ScratchArea](Publisher.ScratchArea.md)** object for a given document.


## Syntax

_expression_.**ScratchArea**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

ScratchArea


## Remarks

The **ScratchArea** object is a collection of objects on the scratch page. The **ScratchArea** object is not in the **Pages** collection because it is fundamentally not a page; its only similarity to a page is that it can contain objects.


## Example

This example sets the variable object as the first shape on the scratch area of the active document.

```vb
Sub ScratchPad() 
 
 Dim saPage As ScratchArea 
 Dim objFirst As Object 
 
 saPage = Application.ActiveDocument.ScratchArea 
 objFirst = saPage.Shapes(1) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]