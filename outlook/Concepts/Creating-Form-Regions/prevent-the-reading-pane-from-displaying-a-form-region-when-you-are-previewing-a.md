---
title: Prevent the Reading Pane from Displaying a Form Region When You are Previewing a Message
ms.prod: outlook
ms.assetid: 46de8d3a-f430-248f-b208-63fee3e9b275
ms.date: 06/08/2019
localization_priority: Normal
---


# Prevent the Reading Pane from Displaying a Form Region When You are Previewing a Message

When you create a form region in a custom form, by default, the form region will be displayed in the Reading Pane when you preview a message that uses that custom form. If you want to prevent the Reading Pane from displaying the form region, specify this in the form region manifest XML file that you register for the form region. For more information on registering a form region, see [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

## To prevent the Reading Pane from displaying a form region

- In the form region manifest XML file, specify **False** as the value of the **showReadingPane** element.

The following example disables the Reading Pane from displaying a form region:

```vb
<showReadingPane>false</showReadingPane>
```

> [!NOTE]
> You can assign **showReadingPane** either a string value or an integer value. The default value is **True** or **1**. To prevent the Reading Pane from displaying the form region, assign either **False** or **0**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]