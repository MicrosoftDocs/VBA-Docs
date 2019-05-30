---
title: Window.Next property (Word)
keywords: vbawd10.chm157417488
f1_keywords:
- vbawd10.chm157417488
ms.prod: word
api_name:
- Word.Window.Next
ms.assetid: 28587dfe-dd49-88b7-0261-b4e42a12eeac
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Next property (Word)

Returns the next document window in the collection of open document windows. Read-only.


## Syntax

_expression_.**Next**

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example activates the next window.

```vb
If Windows.Count > 1 Then ActiveDocument.ActiveWindow.Next.Activate
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]