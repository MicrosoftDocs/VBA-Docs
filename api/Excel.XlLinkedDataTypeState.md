---
title: XlLinkedDataTypeState enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlLinkedDataTypeState
ms.date: 09/12/2018
localization_priority: Normal
---


# XlLinkedDataTypeState enumeration (Excel)

Indicates the state of cells that may contain Linked data types such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877). These are the possible values of the [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md) property.


|Name|Value|Description|
|:-----|:-----|:-----|
| **xlLinkedDataTypeStateNone**|0|The cell does not contain any Linked data types.|
| **xlLinkedDataTypeStateValidLinkedData**|1|The cell contains a Linked data type.|
| **xlLinkedDataTypeStateDisambiguationNeeded**|2|The cell needs to be disambiguated by the user before a Linked data type can be inserted. For example, if the user types "New York" into a cell and attempts to convert it to a "Geography" data type, they may need to select whether they meant New York State or New York City. Until they do so, the cell will be in this state. |
| **xlLinkedDataTypeStateBrokenLinkedData**|3|There is a valid Linked data type in the cell, but entity no longer exists on the service.|
| **xlLinkedDataTypeStateFetchingData**|4|The Linked data type in the cell is in the middle of refreshing new data from the service.|

## See also

- [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]