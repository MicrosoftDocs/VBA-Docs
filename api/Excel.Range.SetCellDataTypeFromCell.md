---
title: Range.SetCellDataTypeFromCell method (Excel)
keywords: vbaxl10.chm144265
f1_keywords:
- vbaxl10.chm144265
ms.prod: excel
api_name:
- Excel.Range.SetCellDataTypeFromCell
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.SetCellDataTypeFromCell method (Excel)

Creates another instance of a Linked data type, such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877), that exists in another cell. The new instance will be linked to the data source in the same way as the original, so it will refresh from the service if you call the **[Workbook.RefreshAll](Excel.Workbook.RefreshAll.md)** method.

## Syntax

_expression_.**SetCellDataTypeFromCell** (_Range_, _LanguageCulture_)

_expression_ A variable that represents the **[Range](excel.range(object).md)** object to which you want to copy the Linked data type.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range _from_ which you want to copy the Linked data type. If the range has more than one cell in it, only the upper-left cell will be used. |
| _LanguageCulture_|Required| **String**|A string representing the [LCID](https://docs.microsoft.com/openspecs/windows_protocols/ms-lcid/a9eac961-e77d-41a6-90a5-ce1a8b0cdb9c) of the language and culture that you would like to use for the linked entity. |

## Example

If you have a Geography Linked data type in cell A1 for the city Seattle, this code copies the Seattle entity to cell B2.

```vb
Range("B2").SetCellDataTypeFromCell Range("A1")
```

After it runs, cells A1 and B2 will contain a Seattle data type, and they will both refresh if you call the **RefreshAll** method. No other cell properties, such as formats, will be copied from A1 to B2.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]