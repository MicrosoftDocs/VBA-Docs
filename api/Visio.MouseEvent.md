---
title: MouseEvent object (Visio)
keywords: vis_sdr.chm52060
f1_keywords:
- vis_sdr.chm52060
ms.prod: visio
api_name:
- Visio.MouseEvent
ms.assetid: 1ae26c28-8fdd-ecfe-b008-d4788c08ce5a
ms.date: 06/19/2019
localization_priority: Normal
---


# MouseEvent object (Visio)

The object passed to **[VisEventProc](visio.iviseventproc.viseventproc.md)** as the subject of **MouseDown**, **MouseMove**, and **MouseUp** events.


## Remarks

The default property of **MouseEvent** is **ToString**. The **ToString** property returns a string that represents the properties of the **MouseEvent** object and has the following form, where _event code_ returns the code of the event that fired (**MouseDown**, **MouseMove**, or **MouseUp**) and **Window.Caption** returns the caption of the window that sourced the event: 

> _event code_; **Button** property value; **KeyButtonState** property value; **x** property value; **y** property value; **Window.Caption**

For example, if a user clicked the left mouse button near the middle of the drawing page while holding down the Shift key, in response to the **MouseDown** event, **ToString** might return `709;1;5;4.3750003+000;4.265000+000;Drawing1`.

Use the **Application** property of the **MouseEvent** object to determine the Microsoft Visio instance hosting the object, and use the **Window** property to determine the Visio window associated with a mouse event.

## Properties

- [Application](Visio.MouseEvent.Application.md)
- [Button](Visio.MouseEvent.Button.md)
- [DragState](Visio.MouseEvent.DragState.md)
- [KeyButtonState](Visio.MouseEvent.KeyButtonState.md)
- [ObjectType](Visio.MouseEvent.ObjectType.md)
- [Stat](Visio.MouseEvent.Stat.md)
- [ToString](Visio.MouseEvent.ToString.md)
- [Window](Visio.MouseEvent.Window.md)
- [x](Visio.MouseEvent.x.md)
- [y](Visio.MouseEvent.y.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]