---
title: Explorer.DisplayMode property (Outlook)
keywords: vbaol11.chm3600
f1_keywords:
- vbaol11.chm3600
ms.assetid: 8e6bcc0d-5a37-2c8f-d059-28706b638dee
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# Explorer.DisplayMode property (Outlook)

Indicates the display mode: Normal, Portrait View, or Portrait Reading Pane.



## Syntax

_expression_.**DisplayMode**

_expression_ A variable that represents an  **[Explorer](Outlook.Explorer.md)** object.


## Modes

• _olDisplayModeNormal_ - This is the normal mode.

• _olDisplayModePortraitView_ - Single pane view. Displays the Portrait View.

• _olDisplayModePortraitReadingPane_ - Single pane view. Displays the Reading Pane.

 **Note** : Outlook is in _olDisplayModeNormal_ when the Reading Pane is turned off. If Outlook is in _olDisplayModeNormal_ and the user turns off the Reading Pane, then Outlook turns off _olDisplayModePortraitView_ mode.


## See also


[Explorer object (Outlook)](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]