---
title: Range.PasteSpecial method (Excel)
keywords: vbaxl10.chm144238
f1_keywords:
- vbaxl10.chm144238
ms.prod: excel
api_name:
- Excel.Range.PasteSpecial
ms.assetid: d3e991f2-7ef7-2ebc-d4bc-ba4c26be472e
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.PasteSpecial method (Excel)

Pastes a **Range** object that has been copied into the specified range.


## Syntax

_expression_.**PasteSpecial** (_Paste_, _Operation_, _SkipBlanks_, _Transpose_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Paste_|Optional| **[XlPasteType](Excel.XlPasteType.md)**| The part of the range to be pasted, such as **xlPasteAll** or **xlPasteValues**.|
| _Operation_|Optional| **[XlPasteSpecialOperation](Excel.XlPasteSpecialOperation.md)**| The paste operation, such as **xlPasteSpecialOperationAdd**.|
| _SkipBlanks_|Optional| **Variant**| **True** to have blank cells in the range on the clipboard not be pasted into the destination range. The default value is **False**.|
| _Transpose_|Optional| **Variant**| **True** to transpose rows and columns when the range is pasted. The default value is **False**.|

## Return value

Variant


## Example

This example replaces the data in cells D1:D5 on Sheet1 with the sum of the existing contents and cells C1:C5 on Sheet1.

```vb
With Worksheets("Sheet1") 
 .Range("C1:C5").Copy 
 .Range("D1:D5").PasteSpecial _ 
  Operation:=xlPasteSpecialOperationAdd 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
