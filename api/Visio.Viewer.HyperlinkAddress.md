---
title: Viewer.HyperlinkAddress property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.HyperlinkAddress
ms.assetid: 13683f2f-3ada-5b45-e9e0-0b2dbed16e94
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.HyperlinkAddress property (Visio Viewer)

Gets the address of the specified hyperlink associated with the specified shape in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**HyperlinkAddress** (_ShapeIndex_, _HyperlinkIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeIndex_|Required| **Long**|The index of the specified shape in the collection of shapes in the drawing that is open in Visio Viewer.|
|_HyperlinkIndex_|Required| **Long**|The index of the specified hyperlink in the collection of hyperlinks in the specified shape in the drawing that is open in Visio Viewer.|

## Return value

**String**


## Remarks

The collections of shapes and hyperlinks are one-based, so the indexes of the first shape in the collection of shapes and the first hyperlink in the collection of hyperlinks are both 1.

The address returned may be a URL or a local file address, depending on the target of the hyperlink. To follow the hyperlink address returned, use the **[FollowHyperlink](Visio.Viewer.FollowHyperlink.md)** method or another link-navigation method exposed in your browser's application programming interface (API).


## Example

The following code shows how to get the address of the first hyperlink in the collection of hyperlinks associated with the first shape in the collection of shapes in the drawing that is open in Visio Viewer.

```vb
Debug.Print Viewer1.HyperlinkAddress(1, 1)
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]