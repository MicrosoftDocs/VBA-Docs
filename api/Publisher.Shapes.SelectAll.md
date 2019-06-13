---
title: Shapes.SelectAll method (Publisher)
keywords: vbapb10.chm2162726
f1_keywords:
- vbapb10.chm2162726
ms.prod: publisher
api_name:
- Publisher.Shapes.SelectAll
ms.assetid: 67b88529-814d-c029-1bde-e5dade87636a
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.SelectAll method (Publisher)

Selects all the shapes in the specified **Shapes** collection.


## Syntax

_expression_.**SelectAll**

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Example

This example selects all the shapes on page one of the active publication.

```vb
ActiveDocument.Pages(1).Shapes.SelectAll
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]