---
title: Shapes.Title property (PowerPoint)
keywords: vbapp10.chm543020
f1_keywords:
- vbapp10.chm543020
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Title
ms.assetid: 61e5f162-d9dd-f8d3-6c15-d5a40c00c10f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Title property (PowerPoint)

Returns a  **[Shape](PowerPoint.Shape.md)** object that represents the slide title. Read-only.


## Syntax

_expression_.**Title**

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Return value

Shape


## Remarks

You can also use the  **[Item](PowerPoint.Placeholders.Item.md)** method of the **[Shapes](PowerPoint.Shapes.md)** or **[Placeholders](PowerPoint.Placeholders.md)** collection to return the slide title.


## Example

This example sets the title text on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!"
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]