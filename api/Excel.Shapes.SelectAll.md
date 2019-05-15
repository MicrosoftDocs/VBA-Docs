---
title: Shapes.SelectAll method (Excel)
keywords: vbaxl10.chm638089
f1_keywords:
- vbaxl10.chm638089
ms.prod: excel
api_name:
- Excel.Shapes.SelectAll
ms.assetid: 322f53c0-3a01-ce08-6112-89447f5ce686
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.SelectAll method (Excel)

Selects all the shapes in the specified **Shapes** collection.


## Syntax

_expression_.**SelectAll**

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Example

This example selects all the shapes on _myDocument_, and then creates a **[ShapeRange](Excel.ShapeRange.md)** collection containing all the shapes.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.SelectAll

Set sr = Selection.ShapeRange 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
