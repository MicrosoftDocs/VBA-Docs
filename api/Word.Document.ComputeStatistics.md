---
title: Document.ComputeStatistics method (Word)
keywords: vbawd10.chm158007414
f1_keywords:
- vbawd10.chm158007414
ms.prod: word
api_name:
- Word.Document.ComputeStatistics
ms.assetid: f6f3c42d-b2c0-f0a8-857f-2a8e314f44fb
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ComputeStatistics method (Word)

Returns a statistic based on the contents of the specified document.  **Long**.


## Syntax

_expression_. `ComputeStatistics`( `_Statistic_` , `_IncludeFootnotesAndEndnotes_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Statistic_|Required| **[WdStatistic](Word.WdStatistic.md)**|The statistic to compute.|
| _IncludeFootnotesAndEndnotes_|Optional| **Variant**| **True** to include footnotes and endnotes when computing statistics. If this argument is omitted, the default value is **False**.|

## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example displays the number of words in the active document, including footnotes.


```vb
MsgBox ActiveDocument.ComputeStatistics(Statistic:=wdStatisticWords, _ 
 IncludeFootnotesAndEndnotes:=True) & " words"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]