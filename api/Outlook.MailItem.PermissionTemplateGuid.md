---
title: MailItem.PermissionTemplateGuid property (Outlook)
keywords: vbaol11.chm3507
f1_keywords:
- vbaol11.chm3507
ms.prod: outlook
api_name:
- Outlook.MailItem.PermissionTemplateGuid
ms.assetid: 33436080-1a1c-dee2-5048-83392c241e86
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.PermissionTemplateGuid property (Outlook)

Returns or sets a  **String** value that represents the GUID of the template file to apply to the **[MailItem](Outlook.MailItem.md)** in order to specify Information Rights Management (IRM) permissions. Read/write.


## Syntax

_expression_. `PermissionTemplateGuid`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property complements the IRM properties on a  **MailItem** object; that is, the **[Permission](Outlook.MailItem.Permission.md)** property and the **[PermissionService](Outlook.MailItem.PermissionService.md)** properties.

In particular, the  **PermissionTemplateGuid** property should be synchronized with the **Permission** property to accurately reflect the permission status of the **MailItem**. Setting the **PermissionTemplateGuid** property to a valid GUID should also incur setting the **Permission** property to **OlPermission.olPermissionTemplate**.

An empty string value for the  **PermissionTemplateGuid** property means that there is no permission template file specified for the **MailItem**. For example, if no IRM has been set up (in which case the **Permission** property is **OlPermission.olUnrestricted**), or the restriction is not to forward the **MailItem** (in which case the **Permission** property is **OlPermission.olDoNotForward**).

If you attempt to set the  **PermissionTemplateGuid** property for a received message (that is, the **[Sent](Outlook.MailItem.Sent.md)** property of the **MailItem** is **True**), Microsoft Outlook returns an error.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]