---
title: LineFormat.Pattern property (PowerPoint)
keywords: vbapp10.chm553011
f1_keywords:
- vbapp10.chm553011
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.Pattern
ms.assetid: 5c4c7e5a-1932-01a4-034d-0a4e98c43174
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.Pattern property (PowerPoint)

Sets or returns a value that represents the pattern applied to the specified line. Read/write.


## Syntax

_expression_.**Pattern**

_expression_ A variable that represents a [LineFormat](PowerPoint.LineFormat.md) object.


## Return value

[MsoPatternType](Office.MsoPatternType.md)


## Example

This example adds a patterned line to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1) 
With myDocument.Shapes.AddLine(10, 100, 250, 0).Line 
    .Weight = 6 
    .ForeColor.RGB = RGB(0, 0, 255) 
    .BackColor.RGB = RGB(128, 0, 0) 
    .Pattern = msoPatternDarkDownwardDiagonal 
End With
```


## See also


[LineFormat Object](PowerPoint.LineFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]