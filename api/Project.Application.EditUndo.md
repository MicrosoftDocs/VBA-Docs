---
title: Application.EditUndo method (Project)
keywords: vbapj.chm201
f1_keywords:
- vbapj.chm201
ms.prod: project-server
api_name:
- Project.Application.EditUndo
ms.assetid: f13ce3a1-f8f2-8b00-d870-6e30f6b772f5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EditUndo method (Project)

Cancels the last user-interface action.


## Syntax

_expression_. `EditUndo`( `_fUndo_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _fUndo_|Optional|**Integer**|Specifies the number of actions to undo. If the total number of actions is less than fUndo,  **EditUndo** undoes all actions.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]