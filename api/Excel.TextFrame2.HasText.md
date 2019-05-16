---
title: TextFrame2.HasText property (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.HasText
ms.assetid: b9c7d9f4-22d3-5a45-e03b-8e06e87a2af9
ms.date: 05/17/2019
localization_priority: Normal
---


# TextFrame2.HasText property (Excel)

Returns whether the specified text frame has text. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**HasText**

_expression_ A variable that represents a **[TextFrame2](Excel.TextFrame2.md)** object.


## Example

This example formats the text in the text frame if the text frame contains text.

```vb
With ActiveSheet.Shapes(1).TextFrame2 
If .HasText Then 
.TextRange2.Font.Name = "Arial" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]