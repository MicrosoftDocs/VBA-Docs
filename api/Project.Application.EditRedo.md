---
title: Application.EditRedo method (Project)
keywords: vbapj.chm200
f1_keywords:
- vbapj.chm200
ms.prod: project-server
api_name:
- Project.Application.EditRedo
ms.assetid: 4d391a2e-cc0b-f2c6-2347-8020ada46670
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EditRedo method (Project)

Redoes the top item on the redo stack.


## Syntax

_expression_. `EditRedo`( `_fRedo_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _fRedo_|Optional|**Integer**|Specifies the number of items to redo. If the total number of items on the redo stack is less than fRedo,  **EditRedo** redoes all items.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]