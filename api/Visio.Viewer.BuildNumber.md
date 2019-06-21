---
title: Viewer.BuildNumber property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.BuildNumber
ms.assetid: 573cc757-5144-77c0-d168-6d8b4c27fe8d
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.BuildNumber property (Visio Viewer)

Gets the build number of Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**BuildNumber**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Long**


## Remarks

For the 2007 release of Visio Viewer, the build number is a four-digit number used by Visio developers for internal purposes.


## Example

The following code gets the build number of Visio Viewer and prints it in the Immediate window.

```vb
Debug.Print vsoViewer.BuildNumber
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]