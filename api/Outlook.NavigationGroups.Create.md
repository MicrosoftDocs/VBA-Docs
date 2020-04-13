---
title: NavigationGroups.Create method (Outlook)
keywords: vbaol11.chm2858
f1_keywords:
- vbaol11.chm2858
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Create
ms.assetid: 5f2bdcfc-4748-4170-7214-bcadc9e3dc36
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups.Create method (Outlook)

Creates and returns a new **[NavigationGroup](Outlook.NavigationGroup.md)** object, added to the end of the **[NavigationGroups](Outlook.NavigationGroups.md)** collection.


## Syntax

_expression_. `Create`( `_GroupDisplayName_` )

_expression_ A variable that represents a [NavigationGroups](Outlook.NavigationGroups.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _GroupDisplayName_|Required| **String**|The value of the  **[Name](Outlook.NavigationGroup.Name.md)** property for the new **NavigationGroup** object.|

## Return value

A **NavigationGroup** object that represents the new navigation group.


## Remarks

A **NavigationGroups** collection can contain multiple **NavigationGroup** objects with the same **Name** property values.

An error occurs if an add-in attempts to add more than 50 navigation groups to a **NavigationGroups** collection, or if an add-in attempts to add a **NavigationGroup** object to the **NavigationGroups** collection of a **[MailModule](Outlook.MailModule.md)** object.


## See also


[NavigationGroups Object](Outlook.NavigationGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]