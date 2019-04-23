---
title: DoCmd.OpenTable method (Access)
keywords: vbaac10.chm4164
f1_keywords:
- vbaac10.chm4164
ms.prod: access
api_name:
- Access.DoCmd.OpenTable
ms.assetid: 6461c8c1-7452-f812-8914-e46406c58eae
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.OpenTable method (Access)

The **OpenTable** method carries out the OpenTable action in Visual Basic.


## Syntax

_expression_.**OpenTable** (_TableName_, _View_, _DataMode_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TableName_|Required|**Variant**|A string expression that's the valid name of a table in the current database. If you execute Visual Basic code containing the **OpenTable** method in a library database, Microsoft Access looks for the table with this name first in the library database, and then in the current database.|
| _View_|Optional|**[AcView](Access.AcView.md)**|An **AcView** constant that specifies the view in which the table will open. The default value is **acViewNormal**.|
| _DataMode_|Optional|**[AcOpenDataMode](Access.AcOpenDataMode.md)**|An **AcOpenDataMode** constant that specifies the data entry mode for the table. The default value is **acEdit**.|

## Remarks

You can use the **OpenTable** method to open a table in Datasheet view, Design view, or Print Preview. You can also select a data entry mode for the table.


## Example

The following example opens the **Employees** table in Print Preview.

```vb
DoCmd.OpenTable "Employees", acViewPreview
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
