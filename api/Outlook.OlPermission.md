---
title: OlPermission enumeration (Outlook)
keywords: vbaol11.chm3101
f1_keywords:
- vbaol11.chm3101
ms.prod: outlook
api_name:
- Outlook.OlPermission
ms.assetid: 11126d37-33da-53f7-f5b6-ea8603998651
ms.date: 06/08/2017
localization_priority: Normal
---


# OlPermission enumeration (Outlook)

Indicates the permission restrictions on an **Item**.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olDoNotForward**|1| **Item** cannot be forwarded.|
| **olPermissionTemplate**|2|Outlook will use an Information Rights Management (IRM) template to determine the access and usage permissions for the item. See  **[MailItem.PermissionService](Outlook.MailItem.PermissionService.md)** and **[SharingItem.PermissionService](Outlook.SharingItem.PermissionService.md)** properties.|
| **olUnrestricted**|0| **Item** has no permission restrictions.|

## Remarks

Used by the [SharingItem.Permission property (Outlook)](Outlook.SharingItem.Permission.md) and [MailItem.Permission property (Outlook)](Outlook.MailItem.Permission.md) to specify the permissions that the recipients will have on the item.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]