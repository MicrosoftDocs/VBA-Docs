---
title: AutoCorrectEntry.Apply method (Word)
keywords: vbawd10.chm155648102
f1_keywords:
- vbawd10.chm155648102
ms.prod: word
api_name:
- Word.AutoCorrectEntry.Apply
ms.assetid: 9427d4a3-e955-7fc6-eec2-d4580e95b13f
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrectEntry.Apply method (Word)

Replaces a range with the value of the specified AutoCorrect entry.


## Syntax

_expression_.**Apply** (_Range_)

_expression_ Required. A variable that represents an '[AutoCorrectEntry](Word.AutoCorrectEntry.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **[Range](Word.Range.md)**|The range to which to apply the AutoCorrect entry.|

## Example

This example adds an AutoCorrect replacement entry, then applies the "sr" AutoCorrect entry to the selected text.


```vb
AutoCorrect.Entries.Add Name:= "sr", Value:= "Stella Richards" 
AutoCorrect.Entries("sr").Apply Selection.Range
```

This example applies the "sr" AutoCorrect entry to the first word in the active document.




```vb
AutoCorrect.Entries("sr").Apply ActiveDocument.Words(1)
```


## See also


[AutoCorrectEntry Object](Word.AutoCorrectEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]