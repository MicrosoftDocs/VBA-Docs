---
title: Shapes.SelectAll method (PowerPoint)
keywords: vbapp10.chm543016
f1_keywords:
- vbapp10.chm543016
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.SelectAll
ms.assetid: 9d3f5b93-2a8b-5b9a-d725-729baa190a38
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.SelectAll method (PowerPoint)

Selects all the shapes in a **[Shapes](PowerPoint.Shapes.md)** collection.


## Syntax

_expression_.**SelectAll**

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Example

This example selects all the shapes on myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.SelectAll
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]