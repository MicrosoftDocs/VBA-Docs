---
title: AccessObject.GetDependencyInfo method (Access)
keywords: vbaac10.chm12756
f1_keywords:
- vbaac10.chm12756
ms.prod: access
api_name:
- Access.AccessObject.GetDependencyInfo
ms.assetid: 33feb9c9-abac-cbe4-acf9-989957f41b7a
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObject.GetDependencyInfo method (Access)

Returns a **[DependencyInfo](Access.DependencyInfo.md)** object that represents the database objects that are dependent upon the specified object.


## Syntax

_expression_.**GetDependencyInfo**

_expression_ A variable that represents an **[AccessObject](Access.AccessObject.md)** object.


## Return value

DependencyInfo


## Remarks

This method will return a run-time error if any of the following conditions are true:

- The **Track name AutoCorrect info** setting (**Tools** menu > **Options** dialog box > **General** tab) is disabled. You can use the following code to enable the **Track name AutoCorrect info** setting and update the dependency information for all of the objects in the database: `Application.SetOption "Track Name AutoCorrect Info", 1`
    
- You have insufficient permissions to check the dependency information for the specified **AccessObject** object.
    
- This method is being called from an Access project (.adp).
    
Access does not search Visual Basic for Applications (VBA) code, macros, or data access pages for dependencies.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]