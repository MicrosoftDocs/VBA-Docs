---
title: NavigationModules object (Outlook)
keywords: vbaol11.chm3192
f1_keywords:
- vbaol11.chm3192
ms.prod: outlook
api_name:
- Outlook.NavigationModules
ms.assetid: 4b0743d3-0a21-488c-27b2-31ae07129a61
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationModules object (Outlook)

Contains a read-only set of  **[NavigationModule](Outlook.NavigationModule.md)** objects that represent the navigation modules displayed by the navigation pane.


## Remarks

Use the  **[Modules](Outlook.NavigationPane.Modules.md)** property of the **[NavigationPane](Outlook.NavigationPane.md)** object to return a **NavigationModules** object.

Use the  **[Item](Outlook.NavigationModules.Item.md)** method to retrieve a **NavigationModule** object by either the name or ordinal position of the navigation module within the collection, or use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method to return a **NavigationModule** object by the navigation module type.

Use the  **[Count](Outlook.NavigationModules.Count.md)** property to return the number of navigation modules contained in the navigation pane.


## Methods



|Name|
|:-----|
|[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)|
|[Item](Outlook.NavigationModules.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.NavigationModules.Application.md)|
|[Class](Outlook.NavigationModules.Class.md)|
|[Count](Outlook.NavigationModules.Count.md)|
|[Parent](Outlook.NavigationModules.Parent.md)|
|[Session](Outlook.NavigationModules.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]