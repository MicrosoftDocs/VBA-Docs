---
title: Viewer.FollowHyperlink method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.FollowHyperlink
ms.assetid: eafbba6d-6429-744a-facd-e3412916a4bf
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.FollowHyperlink method (Visio Viewer)

Follows the hyperlink at the specified index in the specified shape in Microsoft Visio Viewer.


## Syntax

_expression_.**FollowHyperlink** (_ShapeIndex_, _HyperlinkIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeIndex_|Required| **Long**|The index of the shape that contains the hyperlink.|
|_HyperlinkIndex_|Required| **Long**|The index of the hyperlink in the collection of hyperlinks in the specified shape.|

## Return value

Nothing


## Remarks

The collection of hyperlinks is one-based, so the first hyperlink in the collection is at index position 1. If you pass 0 for _HyperlinkIndex_, Visio Viewer navigates to the default hyperlink for the shape, as set in the **Hyperlinks** dialog box (**Insert** menu) in the current Visio document.


## Example

The following code follows the hyperlink in the first index position in the collection of hyperlinks in the first shape on the page in Visio Viewer.

```vb
vsoViewer.FollowHyperlink 1, 1
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]