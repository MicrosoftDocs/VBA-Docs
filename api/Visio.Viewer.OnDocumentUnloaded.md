---
title: Viewer.OnDocumentUnloaded event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnDocumentUnloaded
ms.assetid: b2f1d5ad-122d-6e55-1cb0-63c78f79bc2b
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnDocumentUnloaded event (Visio Viewer)

Occurs after the current document in Microsoft Visio Viewer is unloaded.


## Syntax

_expression_.**OnDocumentUnloaded**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Remarks

You can unload the current document in Visio Viewer programmatically by using the **[Unload](Visio.Viewer.Unload.md)** method.


## Example

```vb
Private Sub vsoViewer_OnDocumentUnloaded()

    Debug.Print "Current document unloaded."

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]