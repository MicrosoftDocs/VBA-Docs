---
title: SharingItem.Permission property (Outlook)
keywords: vbaol11.chm689
f1_keywords:
- vbaol11.chm689
ms.prod: outlook
api_name:
- Outlook.SharingItem.Permission
ms.assetid: fd1ceafe-8c78-8c63-eaf2-aa8cef71a9f3
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Permission property (Outlook)

Sets or returns an **[OlPermission](Outlook.OlPermission.md)** constant that determines what permissions to grant the recipients on the **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `Permission`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The **Permission** property should be synchronized with the **[PermissionTemplateGuid](Outlook.SharingItem.PermissionTemplateGuid.md)** property to accurately reflect the permission status of the **SharingItem**. Setting the **PermissionTemplateGuid** property to a valid GUID should also incur setting the **Permission** property to **OlPermission.olPermissionTemplate**.

 When no Information Rights Management (IRM) has been set up (in which case the **Permission** property is **OlPermission.olUnrestricted**), or the restriction is not to forward the **SharingItem** (in which case the **Permission** property is **OlPermission.olDoNotForward**), the value of the **PermissionTemplateGuid** property should be an empty string.

Although you can view content that is protected by IRM on any computer running the 2007 Microsoft Office system or a later version, you must have Microsoft Office Professional Edition 2003, Microsoft Office Outlook 2007, or a later version of Outlook to create or send an email that is protected by IRM.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]