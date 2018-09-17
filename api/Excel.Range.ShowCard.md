---
title: Range.ShowCard Method (Excel)
keywords: vbaxl10.chm144258
f1_keywords:
- vbaxl10.chm144258
ms.prod: excel
api_name:
- Excel.Range.ShowCard
ms.date: 09/12/2018
---


# Range.ShowCard Method (Excel)

For a cell containing a Linked Data Type such as [Stocks or Geography](https://support.office.com/en-us/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877), this method will cause a card to appear that shows details about the cell (i.e. the same card that the user can view by clicking on the cell icon).

## Syntax

 _expression_. `ShowCard`

 _expression_ A variable that represents a [Range](Excel.Range(Object).md) object.

## Remarks

For ranges of more than one cell, this method will only attempt to show the card for the upper-left cell in the range. If that cell does not contain a Linked Data Type, nothing happens.

## Example

This code will show the card for the Linked Data Type in cell `E5`:

```vb
Range("E5").ShowCard
```

## See also


[Range.ConvertToLinkedDataType](Excel.Range.ConvertToLinkedDataType.md)

[Range.SetCellDataTypeFromCell](Excel.Range.SetCellDataTypeFromCell.md)

[Range.DataTypeToText](Excel.Range.DataTypeToText.md)

[Range.HasRichDataType](Excel.Range.HasRichDataType.md)

[Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md)
