---
title: Application.WebAddToFavorites method (Project)
keywords: vbapj.chm1314
f1_keywords:
- vbapj.chm1314
ms.prod: project-server
api_name:
- Project.Application.WebAddToFavorites
ms.assetid: 3cf8b3e7-4dbf-8555-1662-2412e7d420b0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WebAddToFavorites method (Project)

Adds a link to the current document or selection to the Favorites folder for the user. 


## Syntax

_expression_. `WebAddToFavorites`( `_CurrentLink_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CurrentLink_|Optional|**Boolean**|**True** if a link will be added to the current selection. **False** if a link will be added to the current document. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The Favorites folder is typically  `C:\Users\UserAlias\Favorites`. For a project file named Basic.mpp that is saved in the  `E:\Project\VBA` folder, **WebAddToFavorites** adds a link named Basic that has the following URL: `file:///E:/Project/VBA/Samples/Basic.mpp`

The **WebAddToFavorites** method is unavailable if the file has never been saved.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]