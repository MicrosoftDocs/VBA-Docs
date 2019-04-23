---
title: "How to: Give a Control a 3-D Appearance Programmatically"
keywords: olmain11.chm1045236
f1_keywords:
- olmain11.chm1045236
ms.prod: outlook
ms.assetid: 7e701b10-4b28-aae9-9238-c12fa8e4f885
ms.date: 06/08/2017
localization_priority: Normal
---


# Give a Control a 3-D Appearance Programmatically

The following code example uses the  **[ModifiedFormPages](../../../api/Outlook.Inspector.ModifiedFormPages.md)** property of the current **[Inspector](../../../api/Outlook.Inspector.md)** object to set the **[SpecialEffect](../../../api/Outlook.checkbox.specialeffect.md)** property of a **[CheckBox](../../../api/Outlook.checkbox.md)** on a page named "Test." By setting the **SpecialEffect** property to 2, the **CheckBox** will have a sunken effect.


```vb
Item.GetInspector.ModifiedFormPpages("Test").Checkbox1.SpecialEffect = 2
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]