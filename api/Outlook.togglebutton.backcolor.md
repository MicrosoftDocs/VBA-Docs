---
title: ToggleButton.BackColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 5c3ff165-c5b6-43f2-1f62-398ad5219a43
ms.date: 06/08/2017
localization_priority: Normal
---


# ToggleButton.BackColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the background color of the object. Read/write.


## Syntax

_expression_.**BackColor**

_expression_ A variable that represents a  **ToggleButton** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```vb
RGB(15,200,75)
```

You can only see the background color of an object if the  **[BackStyle](Outlook.togglebutton.backstyle.md)** property is set to 1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]