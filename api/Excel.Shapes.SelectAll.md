---
title: Shapes.SelectAll method (Excel)
keywords: vbaxl10.chm638089
f1_keywords:
- vbaxl10.chm638089
ms.prod: excel
api_name:
- Excel.Shapes.SelectAll
ms.assetid: 322f53c0-3a01-ce08-6112-89447f5ce686
ms.date: 06/08/2017
localization_priority: Priority
---


# Shapes.SelectAll method (Excel)

Selects all the shapes in the specified  **[Shapes](Excel.Shapes.md)** collection.


## Syntax

_expression_. `SelectAll`

_expression_ A variable that represents a [Shapes](./Excel.Shapes.md) object.


## Example

This example selects all the shapes on  `myDocument` and creates a **[ShapeRange](Excel.ShapeRange.md)** collection containing all the shapes.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.SelectAll
```


```vb
Set sr = Selection.ShapeRange 

```


## See also


[Shapes Object](Excel.Shapes.md)

