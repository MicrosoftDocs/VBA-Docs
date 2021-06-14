---
title: Presentation.ReadOnlyRecommended property (PowerPoint)
keywords: vbapp10.chm583136
f1_keywords:
- vbapp10.chm583136
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ReadOnlyRecommended
ms.date: 09/18/2020
ms.author: lindalu
localization_priority: Normal
---


# Presentation.ReadOnlyRecommended property (PowerPoint)

**True** if the presentation was saved as read-only recommended. Read-only **Boolean.**

## Syntax

_expression_.**ReadOnlyRecommended**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.

## Remarks

When you open a presentation that was saved as read-only recommended, Microsoft PowerPoint displays a message recommending that you open the presentation as read-only.

Use the [SaveCopyAs2](PowerPoint.Presentation.SaveCopyAs2.md) method to change this property.

## Example

The following example displays a message indicating if the active presentation is saved as read-only recommended.

```vb
MsgBox "Presentation is saved as read-only recommended: " &
ActivePresentation.ReadOnlyRecommended
```

## See also

[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]