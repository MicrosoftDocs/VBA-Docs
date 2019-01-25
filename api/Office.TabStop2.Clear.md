---
title: TabStop2.Clear method (Office)
ms.prod: office
api_name:
- Office.TabStop2.Clear
ms.assetid: 18087f5f-5886-d349-b002-6830739ff883
ms.date: 01/25/2019
localization_priority: Normal
---


# TabStop2.Clear method (Office)

Removes the specified custom tab stop.


## Syntax

_expression_.**Clear**

_expression_ An expression that returns a **[TabStop2](Office.TabStop2.md)** object.


## Example

This example clears the first custom tab in the first paragraph of the active Microsoft Word document.


```vb
ActiveDocument.Paragraphs(1).TabStops2(1).Clear 

```


## See also

- [TabStop2 object members](overview/Library-Reference/tabstop2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]