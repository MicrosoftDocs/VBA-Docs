---
title: Font object (Microsoft Forms)
keywords: fm20.chm2000520
f1_keywords:
- fm20.chm2000520
ms.prod: office
ms.assetid: f3f05b2d-bb5b-5a6a-a7ad-43fd43934d9e
ms.date: 11/15/2018
localization_priority: Normal
---


# Font object (Microsoft Forms)

Defines the characteristics of the text used by a control or form.

## Remarks

Each control or form has its own **Font** object to let you set its text characteristics independently of the characteristics defined for other controls and forms. Use font properties to specify the font name, to set bold or underlined text, or to adjust the size of the text.

> [!NOTE] 
> The font properties of your form or [container](../../Glossary/vbe-glossary.md#container) determine the default font attributes of controls you put on the form.

The default property for the **Font** object is the **[Name](name-propertye-microsoft-forms.md)** property. If the **Name** property contains a null string, the **Font** object uses the default system font.

## See also

- [Font object](../../../api/outlook.font.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]