---
title: Field.Unlink method (Publisher)
keywords: vbapb10.chm6094857
f1_keywords:
- vbapb10.chm6094857
ms.prod: publisher
api_name:
- Publisher.Field.Unlink
ms.assetid: 4dfe5c29-eb1e-b071-fd86-6ee222455c4e
ms.date: 06/07/2019
localization_priority: Normal
---


# Field.Unlink method (Publisher)

Replaces the specified field or **[Fields](Publisher.Fields.md)** collection with their most recent results.


## Syntax

_expression_.**Unlink**

_expression_ A variable that represents a **[Field](Publisher.Field.md)** object.


## Remarks

When you unlink a field, its current result is converted to text or a graphic and can no longer be updated automatically.


## Example

This example unlinks the first field in shape one on the first page of the active publication.

```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).Unlink
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]