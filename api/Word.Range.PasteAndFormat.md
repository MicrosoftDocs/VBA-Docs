---
title: Range.PasteAndFormat method (Word)
keywords: vbawd10.chm157155740
f1_keywords:
- vbawd10.chm157155740
ms.prod: word
api_name:
- Word.Range.PasteAndFormat
ms.assetid: 39dd8d10-0ab7-10d3-9e48-39a5e342553d
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.PasteAndFormat method (Word)

Pastes the selected table cells and formats them as specified.


## Syntax

_expression_. `PasteAndFormat`( `_Type_` )

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **WdRecoveryType**|The type of formatting to use when pasting the selected table cells.|

## Example

This example pastes a selected Microsoft Excel chart as a picture. This example assumes that the Clipboard contains an Excel chart.


```vb
Sub PasteChart() 
 Selection.PasteAndFormat Type:=wdChartPicture 
End Sub
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]