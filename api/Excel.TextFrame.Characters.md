---
title: TextFrame.Characters method (Excel)
keywords: vbaxl10.chm644078
f1_keywords:
- vbaxl10.chm644078
ms.prod: excel
api_name:
- Excel.TextFrame.Characters
ms.assetid: 20f42207-4d50-1d9f-7dde-c01d7aef0abc
ms.date: 05/17/2019
localization_priority: Normal
---


# TextFrame.Characters method (Excel)

Returns a **[Characters](Excel.Characters.md)** object that represents a range of characters within a shape's text frame. You can use the **Characters** object to add and format characters within the text frame.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

 _expression_ An expression that returns a **[TextFrame](Excel.TextFrame.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either set to 1 or omitted, the **Characters** method returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, the **Characters** method returns the remainder of the string (everything after the character that was set as the _Start_ argument).|

## Return value

**Characters**


## Remarks

The **Characters** object isn't a collection.


## Example

This example formats as bold the third character in the first shape's text frame on the active worksheet.

```vb
With ActiveSheet.Shapes(1).TextFrame 
 .Characters.Text = "abcdefg" 
 .Characters(3, 1).Font.Bold = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
