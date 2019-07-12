---
title: Row.Select method (PowerPoint)
keywords: vbapp10.chm626004
f1_keywords:
- vbapp10.chm626004
ms.prod: powerpoint
api_name:
- PowerPoint.Row.Select
ms.assetid: 67b1c2ce-392e-e257-b722-dfac76e35f24
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Select method (PowerPoint)

Selects the specified object.


## Syntax

_expression_.**Select**

_expression_ A variable that represents a [Row](PowerPoint.Row.md) object.


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


[Row Object](PowerPoint.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]