---
title: NavigationGroups.Item method (Outlook)
keywords: vbaol11.chm2857
f1_keywords:
- vbaol11.chm2857
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Item
ms.assetid: a6521179-fa65-b5af-629a-458a852a29b4
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups.Item method (Outlook)

Returns a  **[NavigationGroup](Outlook.NavigationGroup.md)** object from the collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [NavigationGroups](Outlook.NavigationGroups.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the object.|

## Return value

A  **NavigationGroup** object that represents the specified object.


## Remarks

The index value of a  **NavigationGroup** in the collection represents the ordinal position of the navigation group when displayed in the navigation pane. Changing the position of navigation groups also changes the index values of navigation groups contained within the **[NavigationGroups](Outlook.NavigationGroups.md)** collection.


## See also


[NavigationGroups Object](Outlook.NavigationGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]