---
title: Viewer.OnToolbarCustomized event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnToolbarCustomized
ms.assetid: 02796238-7773-309b-a136-1ded2c09f93f
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnToolbarCustomized event (Visio Viewer)

Occurs when the user customizes the Microsoft Visio Viewer toolbar by adding or removing buttons.


## Syntax

_expression_.**OnToolbarCustomized**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

Nothing


## Remarks

You can customize the toolbar in Visio Viewer by adding or removing buttons. To do so in the user interface, right-click in the toolbar area, and then choose **Customize**. 

You can customize the toolbar programmatically by using the **[ToolbarButtons](Visio.Viewer.ToolbarButtons.md)** property. For the toolbar to be customizable, the **[ToolbarCustomizable](Visio.Viewer.ToolbarCustomizable.md)** property must be set to its default value, **True**.


## Example

The following code shows how to use the **OnToolbarCustomized** event to display a message in the Immediate window when the user customizes the toolbar in Visio Viewer.

```vb
Private Sub vsoViewer_OnToolbarCustomized()

   Debug.Print "The toolbar has been customized!"

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]