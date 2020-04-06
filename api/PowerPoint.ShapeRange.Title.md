---
title: ShapeRange.Title property (PowerPoint)
keywords: vbapp10.chm548097
f1_keywords:
- vbapp10.chm548097
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Title
ms.assetid: bb4e08a3-6517-c500-23ac-ec65b3340f76
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Title property (PowerPoint)

Returns a **[Shape](PowerPoint.Shape.md)** object that represents the slide title. Read-only.


## Syntax

_expression_.**Title**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Remarks

You can also use the  **[Item](PowerPoint.Placeholders.Item.md)** method of the **[Shapes](PowerPoint.Shapes.md)** or **[Placeholders](PowerPoint.Placeholders.md)** collection to return the slide title.


## Example

The following example sets the title text on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!"
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]