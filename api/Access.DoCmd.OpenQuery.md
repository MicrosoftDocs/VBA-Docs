---
title: DoCmd.OpenQuery method (Access)
keywords: vbaac10.chm4162
f1_keywords:
- vbaac10.chm4162
ms.prod: access
api_name:
- Access.DoCmd.OpenQuery
ms.assetid: 3ea20a28-8dd4-e54c-831b-e7e5444aa793
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.OpenQuery method (Access)

The **OpenQuery** method carries out the OpenQuery action in Visual Basic.


## Syntax

_expression_.**OpenQuery** (_QueryName_, _View_, _DataMode_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _QueryName_|Required|**Variant**|A string expression that's the valid name of a query in the current database. If you execute Visual Basic code containing the **OpenQuery** method in a library database, Microsoft Access looks for the query with this name first in the library database, and then in the current database.|
| _View_|Optional|**[AcView](Access.AcView.md)**|An **AcView** constant that specifies the view in which the query will open. The default value is **acViewNormal**.|
| _DataMode_|Optional|**[AcOpenDataMode](Access.AcOpenDataMode.md)**|An **AcOpenDataMode** constant that specifies the data entry mode for the query. The default value is **acEdit**.|

## Remarks

You can use the **OpenQuery** method to open a select or crosstab query in Datasheet view, Design view, or Print Preview. This action runs an action query. You can also select a data entry mode for the query.

> [!NOTE] 
> This method is only available in the Access database environment. See the **OpenView** or **OpenStoredProcedure** methods if you are using the Access Project environment (.adp).

## Example

The following example opens Sales Totals Query in Datasheet view and enables the user to view but not to edit or add records.

```vb
DoCmd.OpenQuery "Sales Totals Query", , acReadOnly
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
