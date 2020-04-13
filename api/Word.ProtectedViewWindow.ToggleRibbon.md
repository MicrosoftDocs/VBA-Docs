---
title: ProtectedViewWindow.ToggleRibbon method (Word)
keywords: vbawd10.chm231735399
f1_keywords:
- vbawd10.chm231735399
ms.prod: word
api_name:
- Word.ProtectedViewWindow.ToggleRibbon
ms.assetid: 767f3efb-2dfe-c202-c544-f09486c660d9
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.ToggleRibbon method (Word)

Shows or hides the ribbon.


## Syntax

_expression_. `ToggleRibbon`

 _expression_ An expression that returns a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Remarks

If the ribbon is visible, the **ToggleRibbon** method hides it; if the ribbon is hidden, the **ToggleRibbon** method shows it.


## Example

The following code example toggles the ribbon for the active Protected View window.


```vb
ActiveProtectedViewWindow.ToggleRibbon
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]