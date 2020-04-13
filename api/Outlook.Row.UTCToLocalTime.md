---
title: Row.UTCToLocalTime method (Outlook)
keywords: vbaol11.chm2247
f1_keywords:
- vbaol11.chm2247
ms.prod: outlook
api_name:
- Outlook.Row.UTCToLocalTime
ms.assetid: 82685689-89af-4c49-1e6b-42e1ecd9d301
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.UTCToLocalTime method (Outlook)

Obtains a **Date** value in a **[Table](Outlook.Table.md)** specified by the **[Row](Outlook.Row.md)** object at _Index_ , that has been converted from Coordinated Universal Time (UTC) to local time.


## Syntax

_expression_. `UTCToLocalTime` (_Index_)

_expression_ A variable that represents a [Row](Outlook.Row.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a **Long** representing the column index for the **[Columns](Outlook.Columns.md)** collection or a **String** representing the **[Name](Outlook.Column.Name.md)** of the **[Column](Outlook.Column.md)**.|

## Return value

A **Date** value that has been converted from a representation in UTC to local time. An error is returned if _Index_ is invalid or the row value indicated by _Index_ is not a **Date** value.


## Remarks

Use the helper functions  **[Row.BinaryToString](Outlook.Row.BinaryToString.md)**, **[Row.LocalTimeToUTC](Outlook.Row.LocalTimeToUTC.md)**, and **Row.UTCToLocalTime** to facilitate type conversion of column values at a specific row.

For information on property value representation in a **Table**, see [Factors Affecting Property Value Representation in the Table and View Classes](../outlook/How-to/Search-and-Filter/factors-affecting-property-value-representation-in-the-table-and-view-classes.md). For information on using Date-time comparisons in  **Table** filters, see [Filtering Items Using a Date-time Comparison](../outlook/How-to/Search-and-Filter/filtering-items-using-a-date-time-comparison.md).


## See also


[Row Object](Outlook.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]