---
title: OutlookBarShortcuts object (Outlook)
keywords: vbaol11.chm3004
f1_keywords:
- vbaol11.chm3004
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts
ms.assetid: 5ee9f085-d2fe-c949-9edc-ad073801ea77
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarShortcuts object (Outlook)

Contains a set of  **[OutlookBarShortcut](Outlook.OutlookBarShortcut.md)** objects representing all shortcuts in a group in the **Shortcuts** pane.


## Remarks

Use the  **[Shortcuts](Outlook.OutlookBarGroup.Shortcuts.md)** property to return the **OutlookBarShortcuts** collection object from the **[OutlookBarGroup](Outlook.OutlookBarGroup.md)** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example shows how to retrieve the  **OutlookBarShortcuts** object.


```vb
Set myShortcuts = myOutlookBarGroup.Shortcuts
```


## Events



|Name|
|:-----|
|[BeforeShortcutAdd](Outlook.OutlookBarShortcuts.BeforeShortcutAdd.md)|
|[BeforeShortcutRemove](Outlook.OutlookBarShortcuts.BeforeShortcutRemove.md)|
|[ShortcutAdd](Outlook.OutlookBarShortcuts.ShortcutAdd.md)|

## Methods



|Name|
|:-----|
|[Add](Outlook.OutlookBarShortcuts.Add.md)|
|[Item](Outlook.OutlookBarShortcuts.Item.md)|
|[Remove](Outlook.OutlookBarShortcuts.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.OutlookBarShortcuts.Application.md)|
|[Class](Outlook.OutlookBarShortcuts.Class.md)|
|[Count](Outlook.OutlookBarShortcuts.Count.md)|
|[Parent](Outlook.OutlookBarShortcuts.Parent.md)|
|[Session](Outlook.OutlookBarShortcuts.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]