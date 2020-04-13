---
title: LineFormat.InsetPen property (Word)
keywords: vbawd10.chm164233330
f1_keywords:
- vbawd10.chm164233330
ms.prod: word
api_name:
- Word.LineFormat.InsetPen
ms.assetid: 6dd5a7b7-bb43-2781-98cc-137537346390
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.InsetPen property (Word)

 **MsoTrue** to draw lines inside a specified shape. Read/write **MsoTriState**.


## Syntax

_expression_.**InsetPen**

_expression_ Required. A variable that represents a **[LineFormat](Word.LineFormat.md)** object.


## Remarks

Use the **InsetPen** property to match up the edges of shapes of equal width but whose line widths vary.


## Example

This example sets all shapes in the active document to draw lines inside the shapes.


```vb
Sub InsetLine() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Shapes 
 shpShape.Line.InsetPen = msoTrue 
 Next shpShape 
End Sub
```


## See also


[LineFormat Object](Word.LineFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]