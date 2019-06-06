---
title: DropCap.Clear method (Publisher)
keywords: vbapb10.chm5505042
f1_keywords:
- vbapb10.chm5505042
ms.prod: publisher
api_name:
- Publisher.DropCap.Clear
ms.assetid: 7c30e774-c520-076a-41d8-7c68679f58bc
ms.date: 06/07/2019
localization_priority: Normal
---


# DropCap.Clear method (Publisher)

Removes the dropped capital letter formatting.


## Syntax

_expression_.**Clear**

_expression_ A variable that represents a **[DropCap](Publisher.DropCap.md)** object.


## Example

This example removes the dropped capital letter formatting in the specified text frame.

```vb
Sub ClearDropCap() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.DropCap.Clear 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]