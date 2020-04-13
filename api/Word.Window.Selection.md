---
title: Window.Selection property (Word)
keywords: vbawd10.chm157417476
f1_keywords:
- vbawd10.chm157417476
ms.prod: word
api_name:
- Word.Window.Selection
ms.assetid: 0e6812cd-8b8a-edaf-cf72-cf899c50f92a
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Selection property (Word)

Returns the **Selection** object that represents a selected range or the insertion point. Read-only.


## Syntax

_expression_.**Selection**

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example copies the selection from window one to the next window.


```vb
If Windows.Count >= 2 Then 
 Windows(1).Selection.Copy 
 Windows(1).Next.Activate 
 Selection.Paste 
End If
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]