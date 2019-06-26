---
title: Windows.OnKeystrokeMessageForAddon event (Visio)
keywords: vis_sdr.chm11750740
f1_keywords:
- vis_sdr.chm11750740
ms.prod: visio
api_name:
- Visio.Windows.OnKeystrokeMessageForAddon
ms.assetid: 4f026d4c-487a-d632-fb70-33d206660571
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.OnKeystrokeMessageForAddon event (Visio)

Occurs when Microsoft Visio receives a keystroke message from Microsoft Windows that is targeted at an add-on window or child of an add-on window.


## Syntax

_expression_.**OnKeystrokeMessageForAddon** (_MSG_)

_expression_ A variable that represents a **[Windows](Visio.Windows.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MSG_|Required| **[IVMSGWRAP]**|The message that Visio receives.|

## Remarks

Returns **True** to indicate that the message was handled by the add-on. Otherwise, returns **False**.

The **OnKeystrokeMessageForAddon** event enables add-ons to intercept and process accelerator and keystroke messages directed at their own add-on windows and child windows of their add-on windows. Only add-on windows created using the **Add** method will source this event.

For this event to fire, the add-on window or one of its child windows must have keystroke focus and the Visio message loop must receive the keystroke message. This event does not fire if the message loop associated with an add-on is handling messages instead of Visio.

Visio fires the **OnKeystrokeMessageForAddon** event when it receives messages in the following range:

|Message|Range|
|:------|:----|
|WM_KEYDOWN|0x0100|
|WM_KEYUP|0x0101|
|WM_CHAR|0x0102|
|WM_DEADCHAR|0x0103|
|WM_SYSKEYDOWN|0x0104|
|WM_SYSKEYUP|0x0105|
|WM_SYSCHAR|0x0106|
|WM_SYSDEADCHAR|0x0107|

The **MSGWrap** object, passed to the event handler when the **OnKeystrokeMessageForAddon** event fires, wraps the Microsoft Windows **MSG** structure, which contains message data. See the **MSGWrap** object for more information, or refer to your Windows documentation.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]