---
title: Application.ResourceSharingPoolAction method (Project)
keywords: vbapj.chm2083
f1_keywords:
- vbapj.chm2083
ms.prod: project-server
api_name:
- Project.Application.ResourceSharingPoolAction
ms.assetid: 0406765b-b6d7-ad6b-c1c2-51bb55591e69
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ResourceSharingPoolAction method (Project)

Performs the specified action on a local resource pool.


## Syntax

_expression_. `ResourceSharingPoolAction`( `_Action_`, `_FileName_`, `_ReadOnly_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _action_|Required|**Long**|The actions to perform on the resource pool. Can be one of the  **[PjPoolAction](Project.PjPoolAction.md)** constants.|
| _FileName_|Optional|**String**|The file name of the resource pool on which to perform the action.|
| _ReadOnly_|Optional|**Boolean**|**True** if the files specified with **FileName** are opened read-only.|

## Return value

 **Boolean**


## Remarks




> [!NOTE] 
> Project Professional can share local resources only when not logged on Project Server. If Project Professional is using a Project Server profile, local resource sharing is unavailable.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]