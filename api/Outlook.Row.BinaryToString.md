---
title: Row.BinaryToString method (Outlook)
keywords: vbaol11.chm2243
f1_keywords:
- vbaol11.chm2243
ms.prod: outlook
api_name:
- Outlook.Row.BinaryToString
ms.assetid: 2416a69f-f0a2-b9a6-6f55-688dcf702824
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.BinaryToString method (Outlook)

Obtains a **String** representing a value that has been converted from a binary value for the parent **[Row](Outlook.Row.md)** at the column specified by _Index_.


## Syntax

_expression_. `BinaryToString` (_Index_)

_expression_ A variable that represents a [Row](Outlook.Row.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a **Long** representing the column index for the **[Columns](Outlook.Columns.md)** collection or a **String** representing the **[Name](Outlook.Column.Name.md)** of the **Column**.|

## Return value

A hexadecimal  **String** value that has been converted from a **PT_BINARY** value for the parent **Row** at the column specified by _Index_. Returns the error, "Cannot convert the column specified by Index to String" if the value specified by _Index_ is not **PT_BINARY**.


## Remarks

Use the helper functions  **Row.BinaryToString**, **[Row.LocalTimeToUTC](Outlook.Row.LocalTimeToUTC.md)**, and **[Row.UTCToLocalTime](Outlook.Row.UTCToLocalTime.md)** to facilitate type conversion of column values at a specific row. For more information on property value representation in a **[Table](Outlook.Table.md)**, see [Factors Affecting Property Value Representation in the Table and View Classes](../outlook/How-to/Search-and-Filter/factors-affecting-property-value-representation-in-the-table-and-view-classes.md).


## See also


[Row Object](Outlook.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]