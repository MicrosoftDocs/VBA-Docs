---
title: OutlookBarShortcut object (Outlook)
keywords: vbaol11.chm337
f1_keywords:
- vbaol11.chm337
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcut
ms.assetid: fae05770-1b06-1ddd-e2db-8428e64bd1e2
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarShortcut object (Outlook)

Represents a shortcut in a group in the  **Shortcuts** pane.


## Remarks

Use the  **[Item](Outlook.OutlookBarShortcuts.Item.md)** method to retrieve the **OutlookBarShortcut** object from an **[OutlookBarShortcuts](Outlook.OutlookBarShortcuts.md)** object. Because the **[Name](Outlook.OutlookBarShortcut.Name.md)** property is the default property of the **OutlookBarShortcut** object, you can identify the shortcut by name.


## Example

The following example retrieves an  **OutlookBarShortcut** object by name.


```vb
Set myOlBarShortcut = myOutlookBarShortcuts.Item("Calendar")
```


## Methods



|Name|
|:-----|
|[SetIcon](Outlook.OutlookBarShortcut.SetIcon.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.OutlookBarShortcut.Application.md)|
|[Class](Outlook.OutlookBarShortcut.Class.md)|
|[Name](Outlook.OutlookBarShortcut.Name.md)|
|[Parent](Outlook.OutlookBarShortcut.Parent.md)|
|[Session](Outlook.OutlookBarShortcut.Session.md)|
|[Target](Outlook.OutlookBarShortcut.Target.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]