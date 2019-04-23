---
title: ReplacementList property (Excel Graph)
keywords: vbagr10.chm3077085
f1_keywords:
- vbagr10.chm3077085
ms.prod: excel
api_name:
- Excel.ReplacementList
ms.assetid: 14209e45-f0e9-a166-7970-ecf3ca79e570
ms.date: 04/12/2019
localization_priority: Normal
---


# ReplacementList property (Excel Graph)

Returns the array of AutoCorrect replacements.

## Syntax

_expression_.**ReplacementList** (_Index_)

_expression_ Required. An expression that returns an **[AutoCorrect](excel.autocorrect-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_ | Optional |**Variant**| The row index of the array of AutoCorrect replacements to be returned. The row is returned as a one-dimensional array with two elements: the first element is the text in column 1, and the second element is the text in column 2.|

## Remarks

Use the **[AddReplacement](Excel.AddReplacement.md)** method to add an entry to the replacement list.


## Example

This example searches the replacement list for Temperature and displays the replacement entry if it exists.

```vb
repl = Application.AutoCorrect.ReplacementList 
For x = 1 To UBound(repl) 
 If repl(x, 1) = "Temperature" Then MsgBox repl(x, 2) 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]