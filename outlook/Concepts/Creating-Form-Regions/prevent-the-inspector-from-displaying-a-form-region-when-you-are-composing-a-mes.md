---
title: Prevent the Inspector from Displaying a Form Region When You are Composing a Message
ms.prod: outlook
ms.assetid: f3162118-9e58-47fb-836e-6b2699bcbd18
ms.date: 06/08/2019
localization_priority: Normal
---


# Prevent the Inspector from Displaying a Form Region When You are Composing a Message

When you create a form region in a custom form, by default, the form region will be displayed in the inspector when you compose a message that uses that custom form. If you want to prevent the inspector from displaying the form region, specify this in the form region manifest XML file that you register for the form region. For more information on registering a form region, see [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

## To prevent the inspector from displaying a form region while in compose mode

- In the form region manifest XML file, specify **False** as the value of the **showInspectorCompose** element.

The following example disables the inspector from displaying a form region when in compose mode:

```vb
<showInspectorCompose>false</showInspectorCompose>
```

> [!NOTE]
> You can assign **showInspectorCompose** either a string value or an integer value. The default value is **True** or **1**. To prevent the inspector from displaying the form region in compose mode, assign either **False** or **0**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]