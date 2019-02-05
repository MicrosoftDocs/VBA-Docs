---
title: Application.CurrentWebUserGroups method (Access)
keywords: vbaac10.chm14600
f1_keywords:
- vbaac10.chm14600
ms.prod: access
api_name:
- Access.Application.CurrentWebUserGroups
ms.assetid: efe80f7a-b6ac-12a5-3704-6e662c87e134
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.CurrentWebUserGroups method (Access)

Gets the collection of Microsoft SharePoint Foundation groups of which the user is a member. 


## Syntax

_expression_.**CurrentWebUserGroups** (_DisplayOption_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DisplayOption_|Required|**[AcWebUserGroupsDisplay](access.acwebusergroupsdisplay.md)**|Specifies the type of information to return about the user's groups.|

## Return value

Variant


## Remarks

The **CurrentWebUserGroups** method returns **Null** if the user is not a member of any groups.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]