---
title: Range.DataTypeToText method (Excel)
keywords: vbaxl10.chm144266
f1_keywords:
- vbaxl10.chm144266
ms.prod: excel
api_name:
- Excel.Range.DataTypeToText
ms.date: 09/12/2018
localization_priority: Normal
---


# Range.DataTypeToText method (Excel)

If any of the cells in the range are a Linked data type such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877), this call will convert their values to text. 

## Syntax

_expression_. `DataTypeToText`

_expression_ A variable that represents a [Range](Excel.Range(Object).md) object.

## Remarks

The call can fail if _none_ of the cells in the range are a Linked data type. In this case, it will throw runtime exception '1004'.

## Example

This code will convert the range E5:G10 into text:

```vb
Worksheets(1).Range("E5:G10").DataTypeToText
```

## See also

- [Range.ConvertToLinkedDataType](Excel.Range.ConvertToLinkedDataType.md)
- [Range.SetCellDataTypeFromCell](Excel.Range.SetCellDataTypeFromCell.md)
- [Range.HasRichDataType](Excel.Range.HasRichDataType.md)
- [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md)
- [Range.ShowCard](Excel.Range.ShowCard.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]