---
title: Viewer.Unload method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.Unload
ms.assetid: 4b746cbf-2f81-b4ef-3f5e-4df93a543292
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.Unload method (Visio Viewer)

Unloads the drawing file that is open in Microsoft Visio Viewer.


## Syntax

_expression_.**Unload**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

Nothing


## Example

The following code prints the name of the drawing that is loaded in Visio Viewer, unloads the drawing, and then prints a blank statement that shows that the document has been unloaded.

```vb
Debug.Print vsoViewer.SRC

vsoViewer.Unload

Debug.Print vsoViewer.SRC
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]