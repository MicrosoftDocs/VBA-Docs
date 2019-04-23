---
title: Frame.Copy Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 61654953-0233-f068-ae50-8f81a51f88d3
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame.Copy Method (Outlook Forms Script)

Copies the contents of an object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a  **Frame** object.


## Remarks

The original content remains on the object.

The actual content that is copied depends on the object. Using  **Copy** for a form, **[Frame](Outlook.frame.md)**, or  **[Page](Outlook.page.md)** copies the currently active control.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]