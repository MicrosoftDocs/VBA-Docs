---
title: Page.Copy Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 6013fe1e-eb1c-dcca-b5eb-d99cc84f22fa
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Copy Method (Outlook Forms Script)

Copies the contents of an object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **Page** object.


## Remarks

The original content remains on the object.

The actual content that is copied depends on the object. Using  **Copy** for a form, **[Frame](Outlook.frame.md)**, or  **[Page](Outlook.page.md)** copies the currently active control.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]