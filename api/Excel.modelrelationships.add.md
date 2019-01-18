---
title: ModelRelationships.Add method (Excel)
keywords: vbaxl10.chm940077
f1_keywords:
- vbaxl10.chm940077
ms.prod: excel
ms.assetid: 9525ce41-1957-cb88-ecdd-9d18295fa422
ms.date: 06/08/2017
localization_priority: Normal
---


# ModelRelationships.Add method (Excel)

Adds a new relationship to the model.


## Syntax

_expression_. `Add`_(ForeignKeyColumn,_ _PrimaryKeyColumn)_

_expression_ A variable that represents a [ModelRelationships object (Excel)](Excel.modelrelationships.md) object (Excel).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ForeignKeyColumn_|Required|MODELTABLECOLUMN|A [ModelTableColumn object (Excel)](Excel.modeltablecolumn.md) object (Excel) representing the foreign key column in the table on the many side of the one-to-many relationship.|
| _PrimaryKeyColumn_|Required|MODELTABLECOLUMN|A [ModelTableColumn object (Excel)](Excel.modeltablecolumn.md) object (Excel) representing the primary key column in the table on the one side of the one-to-many relationship.|

## Return value

 **MODELRELATIONSHIP**


## See also



[ModelRelationships Object](Excel.modelrelationships.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]