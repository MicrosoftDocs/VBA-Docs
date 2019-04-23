---
title: CurrentProject.UpdateDependencyInfo method (Access)
keywords: vbaac10.chm12727
f1_keywords:
- vbaac10.chm12727
ms.prod: access
api_name:
- Access.CurrentProject.UpdateDependencyInfo
ms.assetid: 90461646-22a6-bfa8-4663-9f05c8ac3757
ms.date: 02/27/2019
localization_priority: Normal
---


# CurrentProject.UpdateDependencyInfo method (Access)

Updates the dependency information for the database.


## Syntax

_expression_.**UpdateDependencyInfo**

_expression_ A variable that represents a **[CurrentProject](Access.CurrentProject.md)** object.


## Remarks

The **UpdateDependencyInfo** method opens, saves, and then closes every table, query, form, and report in the database; no messages are presented to the user.

This method will return a run-time error if any of the following conditions are true:

- This method is being called from an Access project (.adp).
    
- Any database objects are open.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]