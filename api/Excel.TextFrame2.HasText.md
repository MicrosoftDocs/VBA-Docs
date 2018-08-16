---
title: TextFrame2.HasText Property (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.HasText
ms.assetid: b9c7d9f4-22d3-5a45-e03b-8e06e87a2af9
ms.date: 06/08/2017
---


# TextFrame2.HasText Property (Excel)

Returns whether the specified text frame has text. Read-only  **[MsoTriState](./Office.MsoTriState.md)** .


## Syntax

 _expression_. `HasText`

 _expression_ A variable that represents a [TextFrame2](./Excel.TextFrame2.md) object.


## Example

This example formats the text in the text frame, if the text frame contains text.


```vb
With ActiveSheet.Shapes(1).TextFrame2 
If .HasText Then 
.TextRange2.Font.Name = "Arial" 

```


## See also


[TextFrame2 Object](Excel.TextFrame2.md)

