---
title: Field.Next property (Publisher)
keywords: vbapb10.chm6094854
f1_keywords:
- vbapb10.chm6094854
ms.prod: publisher
api_name:
- Publisher.Field.Next
ms.assetid: a8f0a246-c55e-715e-3f97-a2f08c383e87
ms.date: 06/07/2019
localization_priority: Normal
---


# Field.Next property (Publisher)

Returns a **Field** object that represents the next field in a text range.


## Syntax

_expression_.**Next**

_expression_ A variable that represents a **[Field](Publisher.Field.md)** object.


## Return value

Field


## Example

This example makes the field next to the first field in the specified text range bold. This assumes that there are at least two fields in the specified text range.

```vb
Sub GoToNextField() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Fields(1).Next.TextRange.Font.Bold = msoTrue 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]