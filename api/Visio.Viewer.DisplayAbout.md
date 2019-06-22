---
title: Viewer.DisplayAbout method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.DisplayAbout
ms.assetid: 53d4e175-4038-94c3-68e3-0a0cb2b8a79a
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.DisplayAbout method (Visio Viewer)

Displays the **About Microsoft Visio Viewer** dialog box in the Microsoft Visio Viewer.


## Syntax

_expression_.**DisplayAbout**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

Nothing


## Remarks

Calling this method is equivalent to choosing the **About Microsoft Visio Viewer** button on the toolbar in the Viewer.


## Example

The following code displays the **About Microsoft Visio Viewer** dialog box.

```vb
vsoViewer.DisplayAbout
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]