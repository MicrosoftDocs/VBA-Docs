---
title: Endnotes.NumberStyle property (Word)
keywords: vbawd10.chm155254885
f1_keywords:
- vbawd10.chm155254885
ms.prod: word
api_name:
- Word.Endnotes.NumberStyle
ms.assetid: 9157acf1-6452-ec85-5032-66cf960b94f4
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.NumberStyle property (Word)

Returns or sets the number style. Read/write  **WdNoteNumberStyle**.


## Syntax

_expression_. `NumberStyle`

 _expression_ An expression that represents a '[Endnotes](Word.endnotes.md)' object.


## Remarks

Some of the constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets the formatting for footnotes and endnotes in the active document.


```vb
With ActiveDocument 
 .Footnotes.NumberStyle = wdNoteNumberStyleLowercaseRoman 
 .Endnotes.NumberStyle = wdNoteNumberStyleUppercaseRoman 
End With
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]