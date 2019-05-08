---
title: Find.Frame property (Word)
keywords: vbawd10.chm162529306
f1_keywords:
- vbawd10.chm162529306
ms.prod: word
api_name:
- Word.Find.Frame
ms.assetid: 66cfee6f-649c-cef9-1dee-d2a4a6de4a7a
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.Frame property (Word)

Returns a  **[Frame](Word.Frame.md)** object that represents the frame formatting for the specified style or find-and-replace operation. Read-only.


## Syntax

_expression_. `Frame`

_expression_ A variable that represents a '[Find](Word.Find.md)' object.


## Example

This example finds the first frame with wrap around formatting. If such a frame is found, a message is displayed on the status bar.


```vb
With ActiveDocument.Content.Find 
 .Text = "" 
 .Frame.TextWrap = True 
 .Execute Forward:=True, Wrap:=wdFindContinue, Format:=True 
 If .Found = True Then StatusBar = "Frame was found" 
 .Parent.Select 
End With
```


## See also


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]