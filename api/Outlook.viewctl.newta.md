---
title: ViewCtl.NewTask Method (Outlook View Control)
ms.prod: outlook
ms.assetid: c997fd53-87fe-11b4-5966-a644bb812332
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewCtl.NewTask Method (Outlook View Control)

Creates and displays a new task.


## Syntax

_expression_.**NewTask**

_expression_ A variable that represents a **ViewCtl** object.


## Remarks

When the new task is saved it is saved to the  **Tasks**folder, if any, that is displayed in the control. If there is no  **Tasks** folder displayed in the control, the task is saved to the user's default **Tasks** folder.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]