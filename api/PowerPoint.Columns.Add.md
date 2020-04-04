---
title: Columns.Add method (PowerPoint)
keywords: vbapp10.chm623004
f1_keywords:
- vbapp10.chm623004
ms.prod: powerpoint
api_name:
- PowerPoint.Columns.Add
ms.assetid: c16d9aa7-20f0-b3f5-e6f2-ad09867d565e
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Add method (PowerPoint)

Adds a new column to an existing table. Returns a **[Column](PowerPoint.Column.md)** object that represents the new table column.


## Syntax

_expression_.**Add** (_BeforeColumn_)

_expression_ A variable that represents a [Columns](PowerPoint.Columns.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BeforeColumn_|Optional|**Long**|The index number that specifies the table column before which the new column will be inserted. |

## Return value

Column


## Remarks

The value of the BeforeColumn parameter must be between 1 and the number of columns in the table, inclusive. The default value is -1, which means that if you omit the BeforeColumn parameter, the new column is added as the last column in the table.


## See also


[Columns Object](PowerPoint.Columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]