---
title: Name property (Microsoft Forms)
keywords: fm20.chm2001590
f1_keywords:
- fm20.chm2001590
ms.prod: office
ms.assetid: e89050ca-a1da-8a05-b47a-774b22dcfa6b
ms.date: 11/16/2018
localization_priority: Normal
---


# Name property (Microsoft Forms)

Specifies the name of a control or an object, or the name of a font to associate with a **[Font](font-object-microsoft-forms.md)** object.

## Syntax

For Font: _Font_.**Name** [= _String_ ]

For all other controls and objects: _object_.**Name** [= _String_ ]

The **Name** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. The name you want to assign to the font or control.|

## Settings

Guidelines for assigning a string to **Name**, such as the maximum length of the name, vary from one application to another.

## Remarks

For objects, the default value of **Name** consists of the object's [class](../../Glossary/vbe-glossary.md#class) name followed by an integer. For example, the default name for the first **[TextBox](textbox-control.md)** you place on a form is TextBox1. The default name for the second **TextBox** is TextBox2.

You can set the **Name** property for a control from the control's property sheet, or for controls added at [run time](../../Glossary/vbe-glossary.md#run-time), by using program statements. If you add a control at [design time](../../Glossary/vbe-glossary.md#design-time), you cannot modify its **Name** property at run time. Each control added to a form at design time must have a unique name.

For **Font** objects, **Name** identifies a particular typeface to use in the text portion of a control, object, or form. The font's appearance on screen and in print may differ, depending on your computer and printer. If you select a font that your system can't display or that isn't installed, the operating system substitutes a similar font.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]