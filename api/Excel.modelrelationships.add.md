---
title: ModelRelationships.Add method (Excel)
keywords: vbaxl10.chm940077
f1_keywords:
- vbaxl10.chm940077
ms.prod: excel
ms.assetid: 9525ce41-1957-cb88-ecdd-9d18295fa422
ms.date: 04/20/2019
localization_priority: Normal
---


# ModelRelationships.Add method (Excel)

Adds a new relationship to the model.


## Syntax

_expression_.**Add** (_ForeignKeyColumn_, _PrimaryKeyColumn_)

_expression_ A variable that represents a **[ModelRelationships](Excel.modelrelationships.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ForeignKeyColumn_|Required|MODELTABLECOLUMN|A **[ModelTableColumn](Excel.modeltablecolumn.md)** object representing the foreign key column in the table on the many side of the one-to-many relationship.|
| _PrimaryKeyColumn_|Required|MODELTABLECOLUMN|A **ModelTableColumn** object representing the primary key column in the table on the one side of the one-to-many relationship.|


## Return value

**MODELRELATIONSHIP**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]