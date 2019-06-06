---
title: Document.ViewTwoPageSpread property (Publisher)
keywords: vbapb10.chm196665
f1_keywords:
- vbapb10.chm196665
ms.prod: publisher
api_name:
- Publisher.Document.ViewTwoPageSpread
ms.assetid: b5e851ff-d5fc-a98d-02b3-7e14c1b957dc
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ViewTwoPageSpread property (Publisher)

Returns **True** if the specified publication should be viewed as a two-page spread. Read/write **Boolean**.


## Syntax

_expression_.**ViewTwoPageSpread**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

Boolean


## Example

This example opens a message box and displays if the current publication should be viewed in the two-page spread mode.

```vb
Sub ViewTwoPage() 
 
 MsgBox "View Two Page Spread = " & _ 
 Application.ActiveDocument.ViewTwoPageSpread 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]