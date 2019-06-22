---
title: Viewer.Src property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.Src
ms.assetid: 1da0ff33-12d6-0102-478d-fae692678c7f
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.Src property (Visio Viewer)

Gets or sets the path to the source file for the drawing in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**Src**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**String**


## Remarks

To produce a viable diagram in Visio Viewer, the source file must be a Visio drawing file (.vsd or .vdx). The file path may be to a URL as well as to a local or networked file.

If the source file is a multipage document, Visio Viewer displays the page that was active the last time the file was saved in Visio, assuming that the file was not subsequently modified outside of Visio. In that case, Visio Viewer displays either the last active page or the first valid page.

If there is no document loaded in Visio Viewer, the **Src** property returns a null string (nothing).


## Example

The following code sets a typical path to a source file in Visio Viewer.

```vb
vsoViewer.Src = "C:\users\Visio User\My Visio Drawing.vsd"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]