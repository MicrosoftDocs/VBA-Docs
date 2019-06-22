---
title: Window.ParentWindow property (Visio)
keywords: vis_sdr.chm11614685
f1_keywords:
- vis_sdr.chm11614685
ms.prod: visio
api_name:
- Visio.Window.ParentWindow
ms.assetid: 923c5f95-8cae-3901-ac03-d8e7668a5b7d
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.ParentWindow property (Visio)

Returns the  **Window** object that is the parent of another **Window** object. Read-only.


## Syntax

_expression_. `ParentWindow`

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Return value

Window


## Remarks

 **ParentWindow** returns nothing and raises no exception if the window is a top-level window. A top-level window is a member of the **Windows** collection of an **Application** object.

Use the  **Parent** property of a **Window** object to get the **Windows** collection to which a **Window** object belongs.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]