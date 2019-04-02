---
title: NavigationGroups.GetDefaultNavigationGroup method (Outlook)
keywords: vbaol11.chm2860
f1_keywords:
- vbaol11.chm2860
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.GetDefaultNavigationGroup
ms.assetid: accdd554-1aa1-b254-7489-67673b889757
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups.GetDefaultNavigationGroup method (Outlook)

Returns the  **[NavigationGroup](Outlook.NavigationGroup.md)** that corresponds to the selected default shared folder group.


## Syntax

_expression_. `GetDefaultNavigationGroup`( `_DefaultFolderGroup_` )

_expression_ A variable that represents a [NavigationGroups](Outlook.NavigationGroups.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DefaultFolderGroup_|Required| **[OlGroupType](Outlook.OlGroupType.md)**|The type of navigation group to be retrieved.|

## Return value

A  **NavigationGroup** object that represents the selected default folder group.


## Remarks

If the default navigation group specified in  _DefaultFolderGroup_ was deleted or otherwise doesn't exist, it is automatically created if the parent **[NavigationModule](Outlook.NavigationModule.md)** object supports the specified navigation group type. An error occurs if the parent **NavigationModule** object does not support creating this navigation group type.


## See also


[NavigationGroups Object](Outlook.NavigationGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]