---
title: Rectangle.Lines property (Word)
keywords: vbawd10.chm234029064
f1_keywords:
- vbawd10.chm234029064
ms.prod: word
api_name:
- Word.Rectangle.Lines
ms.assetid: 00faac63-97a8-8b65-885a-5bfa3729d70e
ms.date: 06/08/2017
localization_priority: Normal
---


# Rectangle.Lines property (Word)

Returns a  **[Lines](Word.Lines.md)** collection that represents the lines in a specified portion of text in a page.


## Syntax

_expression_. `Lines`

 _expression_ An expression that returns a '[Rectangle](Word.Rectangle.md)' object.


## Remarks

Use the **Lines** collection and related objects and properties to programmatically define page layout in a document.


## Example

The following example accesses the collection of lines in the first rectangle in the first page of the active document if the specified rectangle contains text.


```vb
Dim objRectangle As Rectangle 
Dim objLines As Lines 
 
Set objRectangle = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles(1) 
 
If objRectangle.RectangleType = wdTextRectangle Then _ 
 Set objLines = objRectangle.Lines
```


## See also


[Rectangle Object](Word.Rectangle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]