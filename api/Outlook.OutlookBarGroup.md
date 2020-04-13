---
title: OutlookBarGroup object (Outlook)
keywords: vbaol11.chm320
f1_keywords:
- vbaol11.chm320
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroup
ms.assetid: 4ccc4213-5a57-7a8b-4ce5-869a096bd096
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarGroup object (Outlook)

Represents a group of shortcuts in the  **Shortcuts** pane of an explorer window.


## Remarks

Use the  **[Item](Outlook.OutlookBarGroups.Item.md)** method to retrieve the **OutlookBarGroup** object from an **[OutlookBarGroups](Outlook.OutlookBarGroups.md)** object. Because the **[Name](Outlook.OutlookBarGroup.Name.md)** property is the default property of the **OutlookBarGroup** object, you can identify the group by name.


## Example

The following example retrieves an **OutlookBarGroup** object by name.


```vb
Set myOlBarGroup = myOutlookBarGroups.Item("Other Shortcuts")
```


## Properties



|Name|
|:-----|
|[Application](Outlook.OutlookBarGroup.Application.md)|
|[Class](Outlook.OutlookBarGroup.Class.md)|
|[Name](Outlook.OutlookBarGroup.Name.md)|
|[Parent](Outlook.OutlookBarGroup.Parent.md)|
|[Session](Outlook.OutlookBarGroup.Session.md)|
|[Shortcuts](Outlook.OutlookBarGroup.Shortcuts.md)|
|[ViewType](Outlook.OutlookBarGroup.ViewType.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]