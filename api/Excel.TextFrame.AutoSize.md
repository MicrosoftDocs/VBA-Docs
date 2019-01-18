---
title: TextFrame.AutoSize property (Excel)
keywords: vbaxl10.chm644081
f1_keywords:
- vbaxl10.chm644081
ms.prod: excel
api_name:
- Excel.TextFrame.AutoSize
ms.assetid: bf434f76-5749-8163-f737-b3bd624092d5
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.AutoSize property (Excel)

 **True** if the size of the specified object is changed automatically to fit text within its boundaries. Read/write **Boolean**.


## Syntax

_expression_. `AutoSize`

_expression_ A variable that represents a [TextFrame](./Excel.TextFrame.md) object.


## Example

This example adjusts the size of the text frame on shape one to fit its text.


```vb
Worksheets(1).Shapes(1).TextFrame.AutoSize = True
```


## See also


[TextFrame Object](Excel.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]