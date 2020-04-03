---
title: NavigationGroups.SelectedChange event (Outlook)
keywords: vbaol11.chm2913
f1_keywords:
- vbaol11.chm2913
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.SelectedChange
ms.assetid: eb55ed92-1925-9aaa-8fd6-9280cfc8aa47
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups.SelectedChange event (Outlook)

Occurs after the selection state is changed for a navigation folder contained in a  **Calendar** navigation module.


## Syntax

_expression_. `SelectedChange`( `_NavigationFolder_` )

_expression_ A variable that represents a [NavigationGroups](Outlook.NavigationGroups.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NavigationFolder_|Required| **[NavigationFolder](Outlook.NavigationFolder.md)**|The selected navigation folder.|

## Remarks

This event occurs when the selection state changes for a folder in the  **Calendar** navigation module, either by a user checking or unchecking a folder in the **Calendar** navigation module of the navigation pane or by an add-in changing the value of the **[IsSelected](Outlook.NavigationFolder.IsSelected.md)** property for a **NavigationFolder** object contained in the **[NavigationGroups](Outlook.NavigationGroups.md)** collection of a **[CalendarModule](Outlook.CalendarModule.md)** object.


## See also


[NavigationGroups Object](Outlook.NavigationGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]