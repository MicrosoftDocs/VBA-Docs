---
title: Options.PasteAdjustParagraphSpacing property (Word)
keywords: vbawd10.chm162988462
f1_keywords:
- vbawd10.chm162988462
ms.prod: word
api_name:
- Word.Options.PasteAdjustParagraphSpacing
ms.assetid: 0aab4ca9-f453-fdb4-8d2e-f37d1d1dde09
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PasteAdjustParagraphSpacing property (Word)

 **True** if Microsoft Word automatically adjusts the spacing of paragraphs when cutting and pasting selections. Read/write **Boolean**.


## Syntax

_expression_. `PasteAdjustParagraphSpacing`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Example

This example sets Word to automatically adjust the spacing of paragraphs when cutting and pasting selections if the option has been disabled.


```vb
Sub AdjustParaSpace() 
 With Options 
 If .PasteAdjustParagraphSpacing = False Then 
 .PasteAdjustParagraphSpacing = True 
 End If 
 End With 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]