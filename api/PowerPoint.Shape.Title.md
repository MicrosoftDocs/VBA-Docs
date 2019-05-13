---
title: Shape.Title property (PowerPoint)
keywords: vbapp10.chm547088
f1_keywords:
- vbapp10.chm547088
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Title
ms.assetid: fc675bc2-0af9-3f72-9b37-fabd586bbb2d
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Title property (PowerPoint)

Returns a  **[Shape](PowerPoint.Shape.md)** object that represents the slide title. Read-only.


## Syntax

_expression_.**Title**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Remarks

You can also use the  **[Item](PowerPoint.Placeholders.Item.md)** method of the **[Shapes](PowerPoint.Shapes.md)** or **[Placeholders](PowerPoint.Placeholders.md)** collection to return the slide title.


## Example

The following example sets the title text on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!"
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
