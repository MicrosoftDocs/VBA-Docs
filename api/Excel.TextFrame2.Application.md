---
title: TextFrame2.Application property (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.Application
ms.assetid: bb5aeb3a-f8d7-3752-27a5-ff1eedd7d4db
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.Application property (Excel)

Returns an  **[Application](Excel.Application(object).md)** object. Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [TextFrame2](./Excel.TextFrame2.md) object.


## Remarks

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object).


## See also


[TextFrame2 Object](Excel.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]