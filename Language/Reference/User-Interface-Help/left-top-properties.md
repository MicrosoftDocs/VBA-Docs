---
title: Left, Top properties
keywords: fm20.chm916577
f1_keywords:
- fm20.chm916577
ms.prod: office
ms.assetid: 372b97d0-30b9-6def-acac-89416fe8b9fc
ms.date: 11/16/2018
localization_priority: Normal
---


# Left, Top properties

The distance between a control and the left or top edge of the form that contains it.

## Syntax

_object_.**Left** [= _Single_ ] <br/>
_object_.**Top** [= _Single_ ]

The **Left** and **Top** property syntaxes have these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Single_|Optional. A numeric expression specifying the coordinates of an object.|

## Settings

Setting the **Left** or **Top** property to 0 places the control's edge at the left or top edge of its [container](../../Glossary/vbe-glossary.md#container).

## Remarks

For most systems, the recommended range of values for **Left** and **Top** is from -32,767 to +32,767. Other values may also work depending on your system configuration. 

For a **[ComboBox](combobox-control.md)**, values of **Left** and **Top** apply to the text portion of the control, not to the list portion. When you move or size a control, its new **Left** setting is automatically entered in the property sheet. When you print a form, the control's horizontal or vertical location is determined by its **Left** or **Top** setting.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]