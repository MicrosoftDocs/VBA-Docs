---
title: PivotItem.Name property (Excel)
keywords: vbaxl10.chm246078
f1_keywords:
- vbaxl10.chm246078
ms.prod: excel
api_name:
- Excel.PivotItem.Name
ms.assetid: b3861675-1f05-9e0d-442c-1cd95385ca09
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotItem.Name property (Excel)

Returns or sets a  **String** value representing the name of the object.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[PivotItem](Excel.PivotItem.md)** object.


## Remarks

The following table shows example values of the  **Name** property and related properties given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".



|**Property**|**Value (OLAP data source)**|**Value (non-OLAP data source)**|
|:-----|:-----|:-----|
| **[Caption](Excel.PivotItem.Caption.md)**|Paris|Paris|
| **[Name](Excel.PivotItem.Name.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
| **[SourceName](Excel.PivotItem.SourceName.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|(Same as SQL property value, read-only)|
| **[Value](Excel.PivotItem.Value.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|

When specifying an index into the  **[PivotItems](Excel.PivotItems.md)** collection, you can use the syntax shown in the following table.



|**Syntax (OLAP data source)**|**Syntax (non-OLAP data source)**|
|:-----|:-----|
|expression.PivotItems("[Europe].[France].[Paris]")|expression.PivotItems("Paris")|

When using the  **[Item](Excel.IconCriteria.Item.md)** property to reference a specific member of a collection, you can use the text index name as shown in the following table.



|**Name (OLAP data source)**|**Name (non-OLAP data source)**|
|:-----|:-----|
|[Europe].[France].[Paris]|Paris|

## See also


[PivotItem Object](Excel.PivotItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]