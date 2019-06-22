---
title: Viewer.AlertsEnabled property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.AlertsEnabled
ms.assetid: 1bf74608-3652-b015-f862-b503d11e5c77
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.AlertsEnabled property (Visio Viewer)

Gets or sets a value that indicates whether warnings and alerts appear when an error occurs in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**AlertsEnabled**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is for warnings and alerts to appear (**True**).


## Example

The following code shows how to determine whether alerts are enabled in Visio Viewer.

```vb
 Debug.Print vsoViewer.AlertsEnabled
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]