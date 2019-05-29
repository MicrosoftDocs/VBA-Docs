---
title: Window.Activate method (Word)
keywords: vbawd10.chm157417572
f1_keywords:
- vbawd10.chm157417572
ms.prod: word
api_name:
- Word.Window.Activate
ms.assetid: d068e7a1-edb8-b244-a315-be1f92471f4c
ms.date: 08/20/2018
localization_priority: Normal
---


# Window.Activate method (Word)

Activates the specified window.

> [!NOTE] 
> If you are working with a document embedded within another document, this event will not occur.

## Syntax

_expression_.**Activate**

_expression_ Required. A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example activates the next window in the Windows collection.


```vb
Sub NextWindow() 
 'Two or more documents must be open for this statement to execute. 
 ActiveDocument.ActiveWindow.Next.Activate 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]