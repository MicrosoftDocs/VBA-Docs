---
title: ComboBox.BorderColor Property (Outlook Forms Script)
keywords: olfm10.chm2000800
f1_keywords:
- olfm10.chm2000800
ms.prod: outlook
ms.assetid: 53a883aa-e488-a1d9-ef18-7afb1c046869
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.BorderColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the border color of an object. Read/write.


## Syntax

_expression_.**BorderColor**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```vb
RGB(15,200,75)
```

To use the  **BorderColor** property, the **[BorderStyle](Outlook.combobox.borderstyle.md)** property must be set to a value other than 0.

 **BorderStyle** uses **BorderColor** to define the border colors. The **[SpecialEffect](Outlook.combobox.specialeffect.md)** property uses system colors exclusively to define its border colors. For Windows operating systems, system color settings are set using the **Display** icon in **Control Panel**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]