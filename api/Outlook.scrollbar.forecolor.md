---
title: ScrollBar.ForeColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c0e3b0a9-3f0d-4df7-9f27-a44f339a2f47
ms.date: 06/08/2017
localization_priority: Normal
---


# ScrollBar.ForeColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the foreground color of an object. Read/write.


## Syntax

_expression_.**ForeColor**

_expression_ A variable that represents a  **ScrollBar** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```vb
RGB(15,200,75)
```

Use the  **ForeColor** property for controls on forms to make them easy to read or to convey a special meaning. For example, if a text box reports the number of units in stock, you can change the color of the text when the value falls below the reorder level.

The  **ForeColor** sets the color of the arrows.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]