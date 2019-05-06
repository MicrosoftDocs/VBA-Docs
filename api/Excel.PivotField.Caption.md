---
title: PivotField.Caption property (Excel)
keywords: vbaxl10.chm240124
f1_keywords:
- vbaxl10.chm240124
ms.prod: excel
api_name:
- Excel.PivotField.Caption
ms.assetid: 7cd928bf-3f69-0950-5b51-9168192c349e
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.Caption property (Excel)

Returns a **String** value that represents the label text for the pivot field.


## Syntax

_expression_.**Caption**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

The following table shows example values of the **Caption** property and related properties, given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".

|Property|Value (OLAP data source)|Value (non-OLAP data source)|
|:-----|:-----|:-----|
| **Caption**|Paris|Paris|
| **[Name](Excel.PivotField.Name.md)**|[Europe].[France].[Paris] &nbsp; (read-only)|Paris|
| **[SourceName](Excel.PivotField.SourceName.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Same as the SQL property value (read-only)|
| **[Value](Excel.PivotField.Value.md)**|[Europe].[France].[Paris]  &nbsp;(read-only)|Paris|

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