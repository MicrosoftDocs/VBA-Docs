---
title: Range.ConvertToLinkedDataType method (Excel)
keywords: vbaxl10.chm144263
f1_keywords:
- vbaxl10.chm144263
ms.prod: excel
api_name:
- Excel.Range.ConvertToLinkedDataType
ms.date: 09/12/2018
---


# Range.ConvertToLinkedDataType method (Excel)

Attempts to convert all the cells in the range to a Linked data type such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877).

## Syntax

 _expression_. `ConvertToLinkedDataType`( `ServiceID`, `LanguageCulture` )

 _expression_ A variable that represents a [Range](Excel.Range(Object).md) object.


### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ServiceID_|Required| **Long**|The ID of the service that will provide the linked entity.|
| _LanguageCulture_|Required| **String**|A string representing the [LCID](https://msdn.microsoft.com/library/cc233982.aspx) of the language and culture that you would like to use for the linked entity. |

## Remarks

The method will fail in all of these cases:

1. The specified locale is not supported on the specified service.
2. All the cells in the range are blank (that is, there is nothing to convert).
3. All the cells in the range contain a formula. If you want to convert such a range, you need to set the cell values to the current calc result first.
4. The cells in the range have already been converted to the specified data type.

In these cases, the method will throw a runtime exception '1004'.

## Example

This code will convert cell `E5` to a "Stocks" Linked data type in the US-English locale:

```vb
Range("E5").ConvertToLinkedDataType ServiceID:=268435456, LanguageCulture:= "en-US"
```

This code will convert cell E6 to a "Geography" Linked data type in the US-English locale:

```vb
Range("E6").ConvertToLinkedDataType ServiceID:=536870912, LanguageCulture:= "en-US"
```

## See also

- [Range.SetCellDataTypeFromCell](Excel.Range.SetCellDataTypeFromCell.md)
- [Range.DataTypeToText](Excel.Range.DataTypeToText.md)
- [Range.HasRichDataType](Excel.Range.HasRichDataType.md)
- [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md)
- [Range.ShowCard](Excel.Range.ShowCard.md)
