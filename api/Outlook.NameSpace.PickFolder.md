---
title: NameSpace.PickFolder method (Outlook)
keywords: vbaol11.chm768
f1_keywords:
- vbaol11.chm768
ms.prod: outlook
api_name:
- Outlook.NameSpace.PickFolder
ms.assetid: f5c1f35a-8e77-8e7f-fcbe-30c6bc90287a
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.PickFolder method (Outlook)

Displays the  **Pick Folder** dialog box.


## Syntax

_expression_. `PickFolder`

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Return value

A  **[Folder](Outlook.Folder.md)** object that represents the folder that the user selects in the dialog box, or **Nothing** if the dialog box is canceled by the user.


## Remarks

The  **Pick Folder** dialog box is a modal dialog box which means that code execution will not continue until the user either selects a folder or cancels the dialog box.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]