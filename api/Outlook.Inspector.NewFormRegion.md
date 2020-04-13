---
title: Inspector.NewFormRegion method (Outlook)
keywords: vbaol11.chm2981
f1_keywords:
- vbaol11.chm2981
ms.prod: outlook
api_name:
- Outlook.Inspector.NewFormRegion
ms.assetid: a8f3c116-6e90-ce85-d160-2b48798b45cf
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.NewFormRegion method (Outlook)

Opens a new page in design mode in the inspector for a new form region.


## Syntax

_expression_. `NewFormRegion`

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Return value

An **Object** that represents the page displaying the form region in the inspector.


## Remarks

If the inspector is not already in design mode,  **NewFormRegion** will put it in design mode.

This method only opens a page for a new form region in design mode. This method is not intended for creating a form region at runtime.


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]