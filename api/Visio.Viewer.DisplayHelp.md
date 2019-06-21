---
title: Viewer.DisplayHelp method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.DisplayHelp
ms.assetid: 4d31b711-2521-cfd3-7689-0bd8618126b1
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.DisplayHelp method (Visio Viewer)

Displays the Help topic that has the specified ID in Microsoft Visio Viewer.


## Syntax

_expression_.**DisplayHelp** (_TopicID_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_TopicID_|Required| **Long**|The ID of the Help topic to display.|

## Return value

Nothing


## Remarks

The Help topic specified appears in the default browser.


## Example

The following code displays the default Help topic.

```vb
vsoViewer.DisplayHelp(1)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]