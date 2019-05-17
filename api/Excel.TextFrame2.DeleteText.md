---
title: TextFrame2.DeleteText method (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.DeleteText
ms.assetid: e96a305c-085a-d807-1336-9dcc22760a7e
ms.date: 05/17/2019
localization_priority: Normal
---


# TextFrame2.DeleteText method (Excel)

Deletes the text from a text frame and all the associated text properties.


## Syntax

_expression_.**DeleteText**

_expression_ A variable that represents a **[TextFrame2](Excel.TextFrame2.md)** object.


## Remarks

The associated text properties include **Font** attributes such as bold, underline, and so on.


## Example

This example deletes the text in the text frame if the text frame contains text.

```vb
With ActiveSheet.Shapes(1).TextFrame2 
 If .HasText Then 
 .DeleteText ()
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]