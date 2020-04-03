---
title: ListBox.BorderColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 9b004ccd-da97-dd60-9d58-2c9b0db6a26c
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox.BorderColor Property (Outlook Forms Script)

Returns or sets a **Long** that specifies the border color of an object. Read/write.


## Syntax

_expression_.**BorderColor**

_expression_ A variable that represents a **ListBox** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```vb
RGB(15,200,75)
```

To use the  **BorderColor** property, the **[BorderStyle](Outlook.listbox.borderstyle.md)** property must be set to a value other than 0.

 **BorderStyle** uses **BorderColor** to define the border colors. The **[SpecialEffect](Outlook.listbox.specialeffect.md)** property uses system colors exclusively to define its border colors. For Windows operating systems, system color settings are set using the **Display** icon in Control Panel.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]