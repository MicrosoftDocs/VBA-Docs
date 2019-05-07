---
title: PivotItem.Caption property (Excel)
keywords: vbaxl10.chm246090
f1_keywords:
- vbaxl10.chm246090
ms.prod: excel
api_name:
- Excel.PivotItem.Caption
ms.assetid: 5b7f3136-971e-6e11-f709-7fffbc86975a
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItem.Caption property (Excel)

Returns a **String** value that represents the label text for the pivot item.


## Syntax

_expression_.**Caption**

_expression_ A variable that represents a **[PivotItem](Excel.PivotItem.md)** object.


## Remarks

The following table shows example values of the **Caption** property and related properties, given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".

|Property|Value (OLAP data source)|Value (non-OLAP data source)|
|:-----|:-----|:-----|
| **Caption**|Paris|Paris|
| **[Name](Excel.PivotItem.Name.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
| **[SourceName](Excel.PivotItem.SourceName.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Same as the SQL property value (read-only)|
| **[Value](Excel.PivotItem.Value.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|

<br/>

When specifying an index into the **[PivotItems](Excel.PivotItems.md)** collection, you can use the syntax shown in the following table.

|Syntax (OLAP data source)|Syntax (non-OLAP data source)|
|:-----|:-----|
|expression.PivotItems("[Europe].[France].[Paris]")|expression.PivotItems("Paris")|

<br/>

When using the **[Item](Excel.PivotItems.Item.md)** property to reference a specific member of a collection, you can use the text index names shown in the following table.

|Name (OLAP data source)|Name (non-OLAP data source)|
|:-----|:-----|
|[Europe].[France].[Paris]|Paris|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]