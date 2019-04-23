---
title: TextFrame2.Column property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.Column
ms.assetid: a9573a4c-db61-ac40-a931-8e32460d1450
ms.date: 01/25/2019
localization_priority: Normal
---


# TextFrame2.Column property (Office)

Returns the **Column** object that represents the columns of the specified text frame. Read-only.


## Syntax

_expression_.**Column**

_expression_ An expression that returns a **[TextFrame2](Office.TextFrame2.md)** object.


## Example

The following code shows how to set the number of columns in the text frame of the first shape on slide one to 2.

```vb
 ActivePresentation.Slides(1).Shapes(1).TextFrame2.Column.Number = 2
```

## See also

- [TextFrame2 object members](overview/Library-Reference/textframe2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
