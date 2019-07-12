---
title: SlideRange.Select method (PowerPoint)
keywords: vbapp10.chm532011
f1_keywords:
- vbapp10.chm532011
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Select
ms.assetid: bcd521aa-b107-4b49-71d1-86c936077cb1
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Select method (PowerPoint)

Selects the specified object.


## Syntax

_expression_.**Select**

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Remarks

If you try to make a selection that isn't appropriate for the view, your code will fail. For example, you can select a slide in slide sorter view but not in slide view.


## Example

This example selects the first five characters in the title of slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Shapes.Title.TextFrame _
    .TextRange.Characters(1, 5).Select
```

This example selects slide one in the active presentation.




```vb
ActivePresentation.Slides(1).Select
```

This example selects a table that has been added to a new slide in a new presentation. The table has three rows and three columns.




```vb
With Presentations.Add.Slides

    .Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select

End With
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]