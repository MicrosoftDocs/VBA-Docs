---
title: ParagraphFormat.LineSpacing property (Word)
keywords: vbawd10.chm156434541
f1_keywords:
- vbawd10.chm156434541
ms.prod: word
api_name:
- Word.ParagraphFormat.LineSpacing
ms.assetid: 30d067e2-9802-f119-bc4c-bd31bfe187f5
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.LineSpacing property (Word)

Returns or sets the line spacing (in points) for the specified paragraphs. Read/write  **Single**.


## Syntax

_expression_. `LineSpacing`

 _expression_ An expression that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

Use the  **[LinesToPoints](Word.Global.LinesToPoints.md)** method to convert a number of lines to the corresponding value in points. For example, `LinesToPoints(2)` returns the value 24.

The  **LineSpacing** property can be set after the **[LineSpacingRule](Word.ParagraphFormat.LineSpacingRule.md)** property has been set to:


-  **wdLineSpaceAtLeast** the line spacing can be greater than or equal to, but never less than, the specified **LineSpacing** value.
    
-  **wdLineSpaceExactly** the line spacing never changes from the specified **LineSpacing** value, even if a larger font is used within the paragraph.
    
-  **wdLineSpaceMultiple** a **LineSpacing** property value must be specified, in points.
    

## Example

This example sets the line spacing for the selected paragraphs to be at least 24 points.


```vb
With Selection.ParagraphFormat 
 .LineSpacingRule = wdLineSpaceAtLeast 
 .LineSpacing = 24 
End With
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
