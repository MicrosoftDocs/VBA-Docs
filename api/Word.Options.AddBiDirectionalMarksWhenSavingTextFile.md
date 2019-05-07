---
title: Options.AddBiDirectionalMarksWhenSavingTextFile property (Word)
keywords: vbawd10.chm162988440
f1_keywords:
- vbawd10.chm162988440
ms.prod: word
api_name:
- Word.Options.AddBiDirectionalMarksWhenSavingTextFile
ms.assetid: 9a8f5ca0-37eb-ca4d-488c-597f6533d9e4
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AddBiDirectionalMarksWhenSavingTextFile property (Word)

 **True** if Microsoft Word adds bidirectional control characters when saving a document as a text file. Read/write **Boolean**.


## Syntax

_expression_. `AddBiDirectionalMarksWhenSavingTextFile`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Remarks

Saving text files with bidirectional control characters preserves right-to-left and left-to-right properties and the order of neutral characters.


## Example

This example sets Word to add bidirectional control characters when saving a document as a text file.


```vb
Options.AddBiDirectionalMarksWhenSavingTextFile = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]