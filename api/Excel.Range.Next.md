---
title: Range.Next property (Excel)
keywords: vbaxl10.chm144165
f1_keywords:
- vbaxl10.chm144165
ms.prod: excel
api_name:
- Excel.Range.Next
ms.assetid: 10712827-9abd-6b8a-49e5-65e3554fcd87
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Next property (Excel)

Returns a  **[Range](Excel.Range(object).md)** object that represents the next cell.


## Syntax

_expression_. `Next`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Remarks

If the object is a range, this property emulates the TAB key, although the property returns the next cell without selecting it.

On a protected sheet, this property returns the next unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the right of the specified cell.


## See also


[Range Object](Excel.Range(object).md)

