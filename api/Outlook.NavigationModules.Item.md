---
title: NavigationModules.Item method (Outlook)
keywords: vbaol11.chm2800
f1_keywords:
- vbaol11.chm2800
ms.prod: outlook
api_name:
- Outlook.NavigationModules.Item
ms.assetid: ee8fdd9c-2b94-29c3-7622-f6e5c8c5399c
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationModules.Item method (Outlook)

Returns a **[NavigationModule](Outlook.NavigationModule.md)** object from the collection.


## Syntax

_expression_.**Item** (_Index_)

 _expression_ An expression that returns a [NavigationModules](Outlook.NavigationModules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either the integer index of the position of the navigation module in the navigation pane, or the value used to match the default property of an object in the collection.|

## Return value

A **NavigationModule** object that represents the specified object.


## Remarks

The **[Name](Outlook.NavigationModule.Name.md)** property is the default property of the **NavigationModule** object.


## See also


[NavigationModules Object](Outlook.NavigationModules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]