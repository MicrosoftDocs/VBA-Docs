---
title: Change the appearance of a control
keywords: fm20.chm5225237
f1_keywords:
- fm20.chm5225237
ms.prod: office
ms.assetid: b14bb419-dd2f-4f0b-9298-847082d93844
ms.date: 12/29/2018
localization_priority: Normal
---


# Change the appearance of a control

Microsoft Forms includes several properties that let you define the appearance of controls in your application:

- **[ForeColor](../../reference/user-interface-help/forecolor-property-microsoft-forms.md)**   
- **[BackColor](../../reference/user-interface-help/backcolor-property-microsoft-forms.md)**, **[BackStyle](../../reference/user-interface-help/backstyle-property-microsoft-forms.md)**   
- **[BorderColor](../../reference/user-interface-help/bordercolor-property.md)**, **[BorderStyle](../../reference/user-interface-help/borderstyle-property.md)**   
- **[SpecialEffect](../../reference/user-interface-help/specialeffect-property.md)**

**ForeColor** determines the [foreground color](../../Glossary/glossary-vba.md#foreground-color). The foreground color applies to any text associated with the control, such as the caption or the control's contents.

**BackColor** and **BackStyle** apply to the control's background. The background is the area within the control's boundaries, such as the area surrounding the text in a control, but not the control's border. **BackColor** determines the [background color](../../Glossary/glossary-vba.md#background-color). **BackStyle** determines whether the background is [transparent](../../Glossary/glossary-vba.md#transparent). A transparent control background is useful if your application design includes a picture as the main background and you want to see that picture through the control.

**BorderColor**, **BorderStyle**, and **SpecialEffect** apply to the control's border. You can use **BorderStyle** or **SpecialEffect** to choose a type of border. Only one of these two properties can be used at a time. When you assign a value to one of these properties, the system sets the other property to **None**. 

- **SpecialEffect** lets you choose one of several border styles, but only lets you use [system colors](../../Glossary/glossary-vba.md#system-colors) for the border. 

- **BorderStyle** supports only one border style, but lets you choose any color that is a valid setting for **BorderColor**. 

- **BorderColor** specifies the color of the control's border, and is only valid when you use **BorderStyle** to create the border.

## See also

- [Microsoft Forms collections, controls, and objects](../../reference/user-interface-help/objects-microsoft-forms.md)
- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]