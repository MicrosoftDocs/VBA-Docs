---
title: Viewer.GetErrorMessage method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.GetErrorMessage
ms.assetid: 31ede4e5-a7ea-c2b8-784e-2e4c7e8bd9ea
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.GetErrorMessage method (Visio Viewer)

Returns a string that describes the specified error message code in Microsoft Visio Viewer.


## Syntax

_expression_.**GetErrorMessage** (_ErrorCode_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ErrorCode_|Required| **Long**|The error message code for which you want to get a description.|

## Return value

String


## Remarks

If you pass an error code that Visio Viewer does not recognize, the **GetErrorMessage** method returns either a string saying so or nothing.

If you pass the value that the **[LastErrorCode](Visio.Viewer.LastErrorCode.md)** property returns, the **GetErrorMessage** method returns the last error code that Visio Viewer returned.


## Example

The following code shows how to use the **GetErrorMessage** method to get a description of the last error code that Visio Viewer returned.

```vb
Debug.Print vsoViewer.GetErrorMessage(vsoViewer.LastErrorCode)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]