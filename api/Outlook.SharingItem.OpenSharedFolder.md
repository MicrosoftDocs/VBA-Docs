---
title: SharingItem.OpenSharedFolder method (Outlook)
keywords: vbaol11.chm698
f1_keywords:
- vbaol11.chm698
ms.prod: outlook
api_name:
- Outlook.SharingItem.OpenSharedFolder
ms.assetid: 6d365693-8d59-a7a0-d6cb-fe959735d708
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.OpenSharedFolder method (Outlook)

Opens a shared folder offered by a sharing invitation.


## Syntax

_expression_. `OpenSharedFolder`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Return value

A  **[Folder](Outlook.Folder.md)** object that represents the shared folder.


## Remarks

This method allows the recipient of a sharing invitation to open the shared folder offered by the sender. An error occurs if this method is called on a  **[SharingItem](Outlook.SharingItem.md)** object with a **Type** property value other than **olSharingMsgTypeInvite** or **olSharingMsgTypeInviteAndRequest**, or if Outlook cannot connect to the shared folder.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]