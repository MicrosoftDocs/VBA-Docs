---
title: Application.Selection property (Publisher)
keywords: vbapb10.chm131109
f1_keywords:
- vbapb10.chm131109
ms.prod: publisher
api_name:
- Publisher.Application.Selection
ms.assetid: b4a542a7-cb54-476b-9ccf-004ce4b9ec47
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.Selection property (Publisher)

Returns a **[Selection](Publisher.Selection.md)** object that represents a selected range or the cursor.


## Syntax

_expression_.**Selection**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Example

This example tests whether the current selection is text. If it is text, the selected text is then displayed in a message box.

```vb
Sub Selectable() 
 
 If Selection.Type = pbSelectionText Then MsgBox Selection.TextRange 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]