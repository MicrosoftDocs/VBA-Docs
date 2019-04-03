---
title: Page.RedoAction Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: ff5e5487-de4c-0967-a724-e4d2d592ff43
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.RedoAction Method (Outlook Forms Script)

Reverses the effect of the most recent  **Undo** action.


## Syntax

_expression_.**RedoAction**

_expression_ A variable that represents a  **Page** object.


## Return value

A  **Boolean** that specifies **True** if the method succeeds, **False** otherwise.


## Remarks

Redo reverses the last  **Undo**, which is not necessarily the last action taken. Not all actions can be undone.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]