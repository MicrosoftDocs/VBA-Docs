---
title: Range.ComputeStatistics method (Word)
keywords: vbawd10.chm157155506
f1_keywords:
- vbawd10.chm157155506
api_name:
- Word.Range.ComputeStatistics
ms.assetid: 5fbeeffd-f592-3078-cd5b-1e2a90ee5092
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Range.ComputeStatistics method (Word)

Returns a **Long** that represents a statistic based on the contents of the specified range.


## Syntax

_expression_.**ComputeStatistics**( `_Statistic_` )

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Statistic_|Required| **[WdStatistic](Word.WdStatistic.md)**|The type of statistic to compute.|

## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.

The Statistic returned for Table.Cell Range will be inaccurate due to the end-of-cell marker interfering with the Statistic. Truncating the Range to remove the end-of-cell-marker will result in the correct Statistic value.

```vb
Set cell = Documents("Report.doc").Tables(1).Cell(1, 1)
Set myRange = cell.Range
lineCount = myRange.ComputeStatistics(Statistic:=wdStatisticLines) 'Returns 0
rng.MoveEnd wdCharacter, -1 
lineCount = myRange.ComputeStatistics(Statistic:=wdStatisticLines)
MsgBox "The first cell contains " & lineCount _ 
 & " lines."
```

## Example

This example displays the number of words and characters in the first paragraph of Report.doc.


```vb
Set myRange = Documents("Report.doc").Paragraphs(1).Range 
wordCount = myRange.ComputeStatistics(Statistic:=wdStatisticWords) 
charCount = myRange.ComputeStatistics(Statistic:=wdStatisticCharacters) 
MsgBox "The first paragraph contains " & wordCount _ 
 & " words and a total of " & charCount & " characters."
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
