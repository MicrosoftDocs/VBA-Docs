---
title: LetterContent.SenderJobTitle property (Word)
keywords: vbawd10.chm161546363
f1_keywords:
- vbawd10.chm161546363
ms.prod: word
api_name:
- Word.LetterContent.SenderJobTitle
ms.assetid: 6d617773-31b4-084a-0dfd-d539c5f8f6d4
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.SenderJobTitle property (Word)

Returns or sets the job title of the person creating a letter with the Letter Wizard. Read/write  **String**.


## Syntax

 _expression_. `SenderJobTitle`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example retrieves the Letter Wizard elements from the active document and displays the sender's job title.


```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
MsgBox myLetterContent.SenderJobTitle
```


## See also


[LetterContent Object](Word.LetterContent.md)

