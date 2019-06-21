---
title: Viewer.OnDocumentLoaded event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnDocumentLoaded
ms.assetid: e8e8af2e-ae79-052e-fb13-d7aee66e2d0f
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnDocumentLoaded event (Visio Viewer)

Occurs after a document is loaded into Microsoft Visio Viewer.


## Syntax

_expression_.**OnDocumentLoaded**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

Nothing


## Remarks

You can load a document into Visio Viewer programmatically by using the **[Load](Visio.Viewer.Load.md)** method.

To capture the **OnDocumentLoaded** event when you are coding in Visual Basic 6.0, load the document in the **Form_Paint()** procedure. The event may not occur in response to calling the **Load** method within the **Form_Load()** procedure.


## Example

The following code shows how to display a message in the Immediate window when a document is loaded in Visio Viewer, showing the name of the document.

```vb
Private Sub vsoViewer_OnDocumentLoaded()

        Debug.Print "Document loaded is "; vsoViewer.SRC

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]