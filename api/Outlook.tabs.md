---
title: Tabs object (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 0b209e50-60c7-e991-f0fb-627dd17cb7ec
ms.date: 06/08/2017
localization_priority: Normal
---


# Tabs object (Outlook Forms Script)

Represents all the **[Tab](Outlook.tab.md)** controls of a **[TabStrip](Outlook.tabstrip.md)**.


## Remarks

Each **Tabs** collection provides the features to manage the number of tabs in the collection and to identify the tab that is currently in use.

The default value of the **Tabs** collection identifies the current **Tab** of a collection.

You can reference a **Tab** by its index value. The index value reflects the ordinal position of the **Tab** within the collection. The index of the first **Tab** in a collection is 0; the index of the second **Tab** is 1; and so on.


## Methods

|Name|Description|
|:-----|:-----|
| [Add](Outlook.tabs.add.md)|Adds a  [Tab](Outlook.tab.md) to a [Tabs](Outlook.tabs.md) collection.|
| [Clear](Outlook.tabs.clear.md)|Removes all tabs from a **Tabs** collection.|
| [Item](Outlook.tabs.item.md)|Returns a member of a collection, either by position or by name.|
| [Remove](Outlook.tabs.remove.md)|Removes a member from a collection.|


## Properties

|Name|Description|
|:-----|:-----|
| [Count](Outlook.tabs.count.md)|Returns a **Long** that represents the number of objects in a collection. Read-only.|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]