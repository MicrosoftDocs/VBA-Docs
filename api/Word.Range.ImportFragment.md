---
title: Range.ImportFragment method (Word)
keywords: vbawd10.chm157155830
f1_keywords:
- vbawd10.chm157155830
ms.prod: word
api_name:
- Word.Range.ImportFragment
ms.assetid: d9feca50-6370-c1c2-00c0-e64ff7a5adb9
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.ImportFragment method (Word)

Imports a document fragment into the document at the specified range.


## Syntax

_expression_. `ImportFragment`( `_FileName_` , `_MatchDestination_` )

 _expression_ An expression that returns a [Range](./Word.Range.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|Specifies the path and file name where the document fragment is stored.|
| _MatchDestination_|Optional| **Boolean**|Specifies whether to match the destination formatting. If  **False**, the imported document fragment retains the formatting in the original document. Default value is **False**.|

## Return value

Nothing


## Remarks

This method replaces the contents of a range. To stop this from occurring, use the  **[Collapse](Word.Range.Collapse.md)** method before using this method.


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]