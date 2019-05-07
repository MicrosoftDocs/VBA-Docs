---
title: Paragraph.AutoAdjustRightIndent property (Word)
keywords: vbawd10.chm156696700
f1_keywords:
- vbawd10.chm156696700
ms.prod: word
api_name:
- Word.Paragraph.AutoAdjustRightIndent
ms.assetid: 274329db-9c26-e2d2-4fb8-4f7af92b3d83
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.AutoAdjustRightIndent property (Word)

 **True** if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns **wdUndefined** if the **AutoAdjustRightIndent** property is set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `AutoAdjustRightIndent`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets Microsoft Word to automatically adjust the right indent for the selected paragraphs if you've specified a set number of characters per line.


```vb
With Selection.ParagraphFormat 
 .AutoAdjustRightIndent = True 
End With
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]