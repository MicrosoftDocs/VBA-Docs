---
title: PrintOptions.ActivePrinter property (PowerPoint)
keywords: vbapp10.chm517015
f1_keywords:
- vbapp10.chm517015
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.ActivePrinter
ms.assetid: 42a7f4be-f2e6-ccdf-64a9-ef686e8130f1
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintOptions.ActivePrinter property (PowerPoint)

Returns the name of the active printer. Read-only.


## Syntax

_expression_.**ActivePrinter**

_expression_ A variable that represents a [PrintOptions](PowerPoint.PrintOptions.md) object.


## Return value

String


## Remarks

This example displays the name of the active printer.


## Example

This example displays the name of the active printer.


```vb
Public Sub ActivePrinter_Example()

    Debug.Print ActivePresentation.PrintOptions.ActivePrinter

End Sub
```


## See also


[PrintOptions Object](PowerPoint.PrintOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]