---
title: AccessObject.IsDependentUpon method (Access)
keywords: vbaac10.chm12755
f1_keywords:
- vbaac10.chm12755
ms.prod: access
api_name:
- Access.AccessObject.IsDependentUpon
ms.assetid: aba465c5-4176-c69a-8eb8-1a6737b6d8cf
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObject.IsDependentUpon method (Access)

Returns a **Boolean** value that indicates whether the specified object is dependent upon the database object specified in the _ObjectName_ argument.


## Syntax

_expression_.**IsDependentUpon** (_ObjectType_, _ObjectName_)

_expression_ A variable that represents an **[AccessObject](Access.AccessObject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that represents the type of database object to check for dependency.|
| _ObjectName_|Required|**String**|The name of the database object to check for dependency.|

## Return value

Boolean

## Remarks

This method will return a run-time error if any of the following conditions are true:

- The **Track name AutoCorrect info** setting (**Tools** menu > **Options** dialog box > **General** tab) is disabled. You can use the following code to enable the **Track name AutoCorrect info** setting and update the dependency information for all of the objects in the database: `Application.SetOption "Track Name AutoCorrect Info", 1`

- You have insufficient permissions to check the dependency information for the specified **AccessObject** object.

- This method is being called from an Access project (.adp).

Access does not search Visual Basic for Applications (VBA) code, macros, or data access pages for dependencies.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]