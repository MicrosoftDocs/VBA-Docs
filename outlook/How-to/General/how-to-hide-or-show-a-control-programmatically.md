---
title: "How to: Hide or Show a Control Programmatically"
keywords: olmain11.chm1045238
f1_keywords:
- olmain11.chm1045238
ms.prod: outlook
ms.assetid: c6cbadf7-7b10-81de-0abe-65b24c3f46d4
ms.date: 06/08/2019
localization_priority: Normal
---


# Hide or Show a Control Programmatically

The following code example uses the **[ModifiedFormPages](../../../api/Outlook.Inspector.ModifiedFormPages.md)** property of the current **[Inspector](../../../api/Outlook.Inspector.md)** object to set the Microsoft Forms 2.0 **Visible** property of a **[CheckBox](../../../api/Outlook.checkbox.md)** on a page named "Test."


```vb
Item.GetInspector.ModifiedFormPages("Test").Checkbox1.Visible = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]