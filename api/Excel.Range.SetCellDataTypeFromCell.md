---
title: Range.SetCellDataTypeFromCell method (Excel)
keywords: vbaxl10.chm144265
f1_keywords:
- vbaxl10.chm144265
ms.prod: excel
api_name:
- Excel.Range.SetCellDataTypeFromCell
ms.date: 09/12/2018
---


# Range.SetCellDataTypeFromCell method (Excel)

Creates another instance of a Linked data type such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877) that exists in another cell. The new instance will be linked to the data source in the same way as the original, so it will refresh from the service if you call [Workbook.RefreshAll](Excel.Workbook.RefreshAll.md).

## Syntax

 _expression_. `SetCellDataTypeFromCell`( `Range`, `LanguageCulture` )

 _expression_ A variable that represents the [Range](Excel.Range(Object).md) _to_ which you want to copy the Linked data type.


### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range _from_ which you want to copy the Linked data type. If the range has more than one cell in it, only the upper-left cell will be used. |

## Example

If you have a "Geography" Linked data type in cell `A1` for the city "Seattle", this code will copy the "Seattle" entity to cell B2:

```vb
Range("B2").SetCellDataTypeFromCell Range("A1")
```

After it runs, cells `A1` and `B2` will have a "Seattle" data type in them, and they will both refresh if you call [Workbook.RefreshAll](Excel.Workbook.RefreshAll.md). No other cell properties, such as formats, will be copied from `A1` to `B2`.

## See also

- [Range.ConvertToLinkedDataType](Excel.Range.ConvertToLinkedDataType.md)
- [Range.DataTypeToText](Excel.Range.DataTypeToText.md)
- [Range.HasRichDataType](Excel.Range.HasRichDataType.md)
- [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md)
- [Range.ShowCard](Excel.Range.ShowCard.md)
