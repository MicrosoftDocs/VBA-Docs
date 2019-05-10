---
title: Range.HasRichDataType property (Excel)
keywords: vbaxl10.chm144257
f1_keywords:
- vbaxl10.chm144257
ms.prod: excel
api_name:
- Excel.Range.HasRichDataType
ms.date: 09/12/2018
localization_priority: Normal
---


# Range.HasRichDataType property (Excel)

**True** if all cells in the range contain a Rich data type. **False** if none of the cells in the range contains a Rich data type; **null** otherwise. Read-only **Variant**.


## Syntax

_expression_.**HasRichDataType**

_expression_ A variable that represents a [Range](Excel.Range(Object).md) object.


## Remarks

Linked data types such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877) are a kind of Rich data type. For linked types, only cells whose [LinkedDataTypeState property](Excel.Range.LinkedDataTypeState.md) is `ValidLinkedData`, `FetchingData`, or `BrokenLinkedData` will be counted as Rich data types by the `HasRichDataType` property. 

Cells in the `DisambiguationNeeded` or `None` states do _not_ count as Rich data types. See the [XlLinkedDataTypeState enum](Excel.XlLinkedDataTypeState.md) for more information about possible Linked data type states.


## Example

This example prompts the user to select a range on Sheet1. If every cell in the selected range contains a Rich data type, the example displays a message.

```vb
Worksheets("Sheet1").Activate 
Set rr = Application.InputBox( _ 
 prompt:="Select a range on this worksheet", _ 
 Type:=8) 
If rr.HasRichDataType = True Then 
 MsgBox "Every cell in the selection contains a Rich Data" 
End If
```

## See also

- [Range.DataTypeToText](Excel.Range.DataTypeToText.md)
- [Excel.XlLinkedDataTypeState](Excel.XlLinkedDataTypeState.md)
- [Range.ConvertToLinkedDataType](Excel.Range.ConvertToLinkedDataType.md)
- [Range.SetCellDataTypeFromCell](Excel.Range.SetCellDataTypeFromCell.md)
- [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md)
- [Range.ShowCard](Excel.Range.ShowCard.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]