---
title: Selection.PasteAndFormat method (Word)
keywords: vbawd10.chm158663669
f1_keywords:
- vbawd10.chm158663669
ms.prod: word
api_name:
- Word.Selection.PasteAndFormat
ms.assetid: 7ed87209-b786-280e-f3f0-dd81eda6f82d
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.PasteAndFormat method (Word)

Pastes the selected table cells and formats them as specified.


## Syntax

_expression_. `PasteAndFormat`( `_Type_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[WdRecoveryType](Word.WdRecoveryType.md)**|The type of formatting to use when pasting the selected table cells.|

## Example

This example pastes a selected Microsoft Excel chart as a picture. This example assumes that the Clipboard contains an Excel chart.


```vb
Sub PasteChart() 
 Selection.PasteAndFormat Type:=wdChartPicture 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
