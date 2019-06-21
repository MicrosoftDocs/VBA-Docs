---
title: Viewer.LastErrorCode property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.LastErrorCode
ms.assetid: cbef3230-128c-3976-04da-eec6da9f6225
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.LastErrorCode property (Visio Viewer)

Gets the last error code returned by Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**LastErrorCode**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Long**


## Remarks

The default value of the **LastErrorCode** property is 0.

To determine what the error code returned by the **LastErrorCode** property means, you can pass the code to the **[GetErrorMessage](Visio.Viewer.GetErrorMessage.md)** method.


## Example

The following code gets the last error code returned by Visio Viewer.

```vb
Debug.Print vsoViewer.LastErrorCode
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]