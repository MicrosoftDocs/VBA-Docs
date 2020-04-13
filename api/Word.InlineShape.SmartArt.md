---
title: InlineShape.SmartArt property (Word)
keywords: vbawd10.chm162005148
f1_keywords:
- vbawd10.chm162005148
ms.prod: word
api_name:
- Word.InlineShape.SmartArt
ms.assetid: fbc47fec-04c4-108c-3280-0931f77b4cb5
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.SmartArt property (Word)

Returns a [SmartArt](Office.SmartArt.md) object that provides a way to work with the SmartArt associated with the specified inline shape. Read-only.


## Syntax

_expression_.**SmartArt**

_expression_ A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Remarks

The **SmartArt** property provides an entry point for interacting with a SmartArt graphic associated with the inline shape.


## Example

The following code example adds a SmartArt graphic to the active document.


```vb
Dim myDoc As Document 
Dim myInlineShape As InlineShape 
Dim mySmartArt As SmartArt 
 
Set myDoc = ActiveDocument 
Set myInlineShape = myDoc.InlineShapes.AddSmartArt(Application.SmartArtLayouts(2), myDoc.Paragraphs(2).Range) 
Set mySmartArt = myShape.SmartArt 

```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]