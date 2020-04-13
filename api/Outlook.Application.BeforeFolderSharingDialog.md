---
title: Application.BeforeFolderSharingDialog event (Outlook)
keywords: vbaol11.chm447
f1_keywords:
- vbaol11.chm447
ms.prod: outlook
api_name:
- Outlook.Application.BeforeFolderSharingDialog
ms.assetid: e06257eb-f2d9-63cf-1220-dda55ee0ea14
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BeforeFolderSharingDialog event (Outlook)

Occurs before the  **Sharing** dialog box is displayed for a selected **[Folder](Outlook.Folder.md)** object.


## Syntax

_expression_. `BeforeFolderSharingDialog`( `_FolderToShare_` , `_Cancel_` )

 _expression_ An expression that returns an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FolderToShare_|Required| **Folder**|The **Folder** object to be shared.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the dialog box is not displayed.|

## Remarks

This event provides an add-in with the capability of replacing the sharing user interface supplied by Outlook with a custom user interface. This event does not occur if a sharing message is programmatically created and displayed.


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]