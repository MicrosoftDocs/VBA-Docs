---
title: Viewer.Load method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.Load
ms.assetid: 8d33e759-793c-2e3c-3731-131fd51b415a
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.Load method (Visio Viewer)

Loads a drawing file into Microsoft Visio Viewer.


## Syntax

_expression_.**Load** (_UrlOrFileName_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_UrlOrFileName_|Required| **String**|The full path and file name or the URL of the file to be loaded.|

## Return value

**Boolean**


## Remarks

If the load succeeds, the **Load** method returns **True**. The method returns **False** if the load fails.

To produce a viable diagram in Visio Viewer, the source file loaded must be a Visio drawing file (.vsd or .vdx). The file path may be to a URL as well as to a local or networked file.

If the source file is a multipage document, Visio Viewer displays the page that was active the last time the file was saved in Visio, assuming that the file was not subsequently modified outside of Visio. In that case, Visio Viewer displays either the last active page or the first valid page.


## Example

The following code loads a drawing named Shapes.vsd from the local drive into Visio Viewer and returns whether the load was successful.

```vb
vsoViewer.Load "C:\Users\User\Shapes.vsd"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]