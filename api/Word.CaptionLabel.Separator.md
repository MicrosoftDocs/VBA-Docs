---
title: CaptionLabel.Separator property (Word)
keywords: vbawd10.chm158924806
f1_keywords:
- vbawd10.chm158924806
ms.prod: word
api_name:
- Word.CaptionLabel.Separator
ms.assetid: b49e1c5d-737e-2084-ec33-71c3a0fa58bc
ms.date: 06/08/2017
localization_priority: Normal
---


# CaptionLabel.Separator property (Word)

Returns or sets the character between the chapter number and the sequence number. Read/write  **WdSeparatorType**.


## Syntax

_expression_.**Separator**

_expression_ Required. A variable that represents a '[CaptionLabel](Word.CaptionLabel.md)' object.


## Example

This example inserts a Figure caption that has a colon (:) between the chapter number and the sequence number.


```vb
With CaptionLabels("Figure") 
 .Separator = wdSeparatorColon 
 .IncludeChapterNumber = True 
End With 
Selection.InsertCaption "Figure"
```


## See also


[CaptionLabel Object](Word.CaptionLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]