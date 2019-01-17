---
title: Viewer.LastErrorCode Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.LastErrorCode
ms.assetid: cbef3230-128c-3976-04da-eec6da9f6225
ms.date: 06/08/2017
localization_priority: Normal
---


# Viewer.LastErrorCode Property (Visio Viewer)

Gets the last error code returned by Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **LastErrorCode**

 _expression_An expression that returns a  **Viewer** object.


## Return value

 **Long**


## Remarks

The default value of the  **LastErrorCode** property is 0.

To determine what the error code returned by the  **LastErrorCode** property means, you can pass the code to the **[GetErrorMessage](Visio.GetErrorMessage.md)** method.


## Example

The following code gets the last error code returned by Visio Viewer.


```vb
Debug.Print vsoViewer.LastErrorCode
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]