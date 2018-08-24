---
title: Range.PasteSpecial Method (Excel)
keywords: vbaxl10.chm144238
f1_keywords:
- vbaxl10.chm144238
ms.prod: excel
api_name:
- Excel.Range.PasteSpecial
ms.assetid: d3e991f2-7ef7-2ebc-d4bc-ba4c26be472e
ms.date: 06/08/2017
---


# Range.PasteSpecial Method (Excel)

Pastes a  **[Range](Excel.Range(object).md)** that has been copied into the specified range.


## Syntax

 _expression_. `PasteSpecial`( `_Paste_` , `_Operation_` , `_SkipBlanks_` , `_Transpose_` )

 _expression_ A variable that represents a [Range](./Excel.Range(Graph property).md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Paste_|Optional| **[XlPasteType](Excel.XlPasteType.md)**| The part of the range to be pasted, such as **xlPasteAll** or **xlPasteValues**.|
| _Operation_|Optional| **[XlPasteSpecialOperation](Excel.XlPasteSpecialOperation.md)**| The paste operation, such as **xlPasteSpecialOperationAdd**.|
| _SkipBlanks_|Optional| **Variant**| **True** to have blank cells in the range on the clipboard not be pasted into the destination range. The default value is **False**.|
| _Transpose_|Optional| **Variant**| **True** to transpose rows and columns when the range is pasted.The default value is **False**.|

### Return Value

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


## See also


[Range Object](Excel.Range(object).md)

