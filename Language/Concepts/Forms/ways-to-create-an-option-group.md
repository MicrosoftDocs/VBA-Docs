---
title: Create an option group
keywords: fm20.chm5225256
f1_keywords:
- fm20.chm5225256
ms.prod: office
ms.assetid: 03e01236-e877-11a1-5de7-52d6307185e7
ms.date: 12/29/2018
localization_priority: Normal
---


# Create an option group

By default, all **[OptionButton](../../reference/user-interface-help/optionbutton-control.md)** controls on a [container](../../Glossary/vbe-glossary.md#container) (such as a form, a **[MultiPage](../../reference/user-interface-help/multipage-control.md)**, or a **[Frame](../../reference/user-interface-help/frame-control.md)**) are part of a single option group. This means that selecting one of the buttons automatically sets all other option buttons on the form to **False**.

If you want more than one option group on the form, there are two ways to create additional groups:

- Use the **[GroupName](../../reference/user-interface-help/groupname-property.md)** property to identify related buttons.
    
- Put related buttons in a **[Frame](../../reference/user-interface-help/frame-control.md)** on the form.
    
The first method is recommended over the second because it reduces the number of controls required in the application. This reduces the disk space required for your application and can improve the performance of your application as well.

> [!NOTE] 
> A **[TabStrip](../../reference/user-interface-help/tabstrip-control.md)** is not a container. Option buttons in the **TabStrip** are included in the form's option group. You can use **GroupName** to create an option group from buttons in a **TabStrip**.

## See also

- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]