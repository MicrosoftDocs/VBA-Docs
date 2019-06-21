---
title: ValidationIssue.TargetPage property (Visio)
keywords: vis_sdr.chm18662675
f1_keywords:
- vis_sdr.chm18662675
ms.prod: visio
api_name:
- Visio.ValidationIssue.TargetPage
ms.assetid: 30aa5d13-93ad-cf55-08ee-9c8b387d6f25
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationIssue.TargetPage property (Visio)

Returns the  **[Page](Visio.Page.md)** object that is associated with the validation issue. Read-only.


## Syntax

_expression_. `TargetPage`

_expression_ A variable that represents a **[ValidationIssue](Visio.ValidationIssue.md)** object.


## Return value

 **Page**


## Remarks

If the issue is not associated with a specific page, the  **TargetPage** property returns **Nothing**.

If the target page is not valid (for example, if it has been deleted), the  **TargetPage** property returns an error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]