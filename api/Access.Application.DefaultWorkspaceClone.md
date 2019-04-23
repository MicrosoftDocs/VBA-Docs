---
title: Application.DefaultWorkspaceClone method (Access)
keywords: vbaac10.chm12550
f1_keywords:
- vbaac10.chm12550
ms.prod: access
api_name:
- Access.Application.DefaultWorkspaceClone
ms.assetid: f72522e5-dd8d-2cd1-df40-4457ef7f94a6
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.DefaultWorkspaceClone method (Access)

You can use the **DefaultWorkspaceClone** method to create a new **Workspace** object without requiring the user to sign in again. For example, if you need to conduct two sets of transactions simultaneously in separate workspaces, you can use the **DefaultWorkspaceClone** method to create a second **Workspace** object with the same user name and password without prompting the user for this information again.


## Syntax

_expression_.**DefaultWorkspaceClone**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Return value

Workspace


## Remarks

> [!NOTE] 
> In Microsoft Access, the **DefaultWorkspaceClone** method is included in this version of Microsoft Access only for compatibility with previous versions using Data Access Object (DAO) language.

The **DefaultWorkspaceClone** method creates a clone of the default **Workspace** object in Microsoft Access. The properties of the **Workspace** object clone have settings identical to those of the default **Workspace** object, except for the **Name** property setting. For the default **Workspace** object, the value of the **Name** property is always #Default Workspace#. For the cloned **Workspace** object, it is #CloneAccess#.

The **UserName** property of the default **Workspace** object indicates the name under which the current user signed in. The **Workspace** object clone is equivalent to the **Workspace** object that would be created if the same user signed in again with the same password.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]