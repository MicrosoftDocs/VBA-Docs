---
title: Index.AccentedLetters property (Word)
keywords: vbawd10.chm159186951
f1_keywords:
- vbawd10.chm159186951
ms.prod: word
api_name:
- Word.Index.AccentedLetters
ms.assetid: 7358af59-a4ee-e509-2a46-d5499dc680d0
ms.date: 06/08/2017
localization_priority: Normal
---


# Index.AccentedLetters property (Word)

 **True** if the specified index contains separate headings for accented letters (for example, words that begin with "À" are under one heading and words that begin with "A" are under another). Read/write **Boolean**.


## Syntax

_expression_. `AccentedLetters`

_expression_ A variable that represents a '[Index](Word.Index.md)' object.


## Example

This example formats the first index in the active document in a single column, with the appropriate letter preceding each alphabetical group and separate headings for accented letters.


```vb
If ActiveDocument.Indexes.Count >= 1 Then 
 With ActiveDocument.Indexes(1) 
 .HeadingSeparator = wdHeadingSeparatorLetter 
 .NumberOfColumns = 1 
 .AccentedLetters = True 
 End With 
End If
```


## See also


[Index Object](Word.Index.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]