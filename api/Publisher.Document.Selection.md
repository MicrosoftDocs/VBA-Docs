---
title: Document.Selection Property (Publisher)
keywords: vbapb10.chm196658
f1_keywords:
- vbapb10.chm196658
ms.prod: publisher
api_name:
- Publisher.Document.Selection
ms.assetid: b1098cdb-8fb7-0906-b193-6dc572ac2993
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Selection Property (Publisher)

Returns a  **[Selection](Publisher.Selection.md)** object that represents a selected range or the cursor.


## Syntax

 _expression_. **Selection**

 _expression_ A variable that represents a  **Document** object.


## Example

This example tests whether the current selection is text. If it is text, the selected text is then displayed in a message box.


```vb
Sub Selectable() 
 
 If Selection.Type = pbSelectionText Then MsgBox Selection.TextRange 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]