---
title: OutlookBarGroups object (Outlook)
keywords: vbaol11.chm3002
f1_keywords:
- vbaol11.chm3002
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups
ms.assetid: bb5fef46-b15a-51c3-0adf-f94e9da6c921
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarGroups object (Outlook)

Contains a set of  **[OutlookBarGroup](Outlook.OutlookBarGroup.md)** objects representing all groups in the Outlook Bar.


## Remarks

Use the  **[Groups](Outlook.OutlookBarStorage.Groups.md)** property to return the **OutlookBarGroups** object from the **[OutlookBarStorage](Outlook.OutlookBarStorage.md)** object.

Use  **Groups** (_index_), where _index_ is the name of an available group, to return a single **OutlookBarGroup** object.


## Example

The following Visual Basic for Applications (VBA) example retrieves the  **OutlookBarGroups** collection from an **OutlookBarStorage** object.


```vb
Set myGroups = myOutlookBarStorage.Groups
```


## Events



|Name|
|:-----|
|[BeforeGroupAdd](Outlook.OutlookBarGroups.BeforeGroupAdd.md)|
|[BeforeGroupRemove](Outlook.OutlookBarGroups.BeforeGroupRemove.md)|
|[GroupAdd](Outlook.OutlookBarGroups.GroupAdd.md)|

## Methods



|Name|
|:-----|
|[Add](Outlook.OutlookBarGroups.Add.md)|
|[Item](Outlook.OutlookBarGroups.Item.md)|
|[Remove](Outlook.OutlookBarGroups.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.OutlookBarGroups.Application.md)|
|[Class](Outlook.OutlookBarGroups.Class.md)|
|[Count](Outlook.OutlookBarGroups.Count.md)|
|[Parent](Outlook.OutlookBarGroups.Parent.md)|
|[Session](Outlook.OutlookBarGroups.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]