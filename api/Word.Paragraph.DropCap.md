---
title: Paragraph.DropCap property (Word)
keywords: vbawd10.chm156696589
f1_keywords:
- vbawd10.chm156696589
ms.prod: word
api_name:
- Word.Paragraph.DropCap
ms.assetid: 32fb0f84-bef9-13cc-86e3-4f644cb76bc7
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.DropCap property (Word)

Returns a  **[DropCap](Word.DropCap.md)** object that represents a dropped capital letter for the specified paragraph. Read-only.


## Syntax

_expression_. `DropCap`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets a dropped capital letter for the first paragraph in the active document.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .FontName = "Arial" 
 .Position = wdDropNormal 
 .LinesToDrop = 3 
 .DistanceFromText = InchesToPoints(0.1) 
End With
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]