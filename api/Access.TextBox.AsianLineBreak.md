---
title: TextBox.AsianLineBreak property (Access)
keywords: vbaac10.chm11147
f1_keywords:
- vbaac10.chm11147
ms.prod: access
api_name:
- Access.TextBox.AsianLineBreak
ms.assetid: 2ee42bb4-e6ae-c6b4-ef6a-71de5d35edad
ms.date: 03/26/2019
localization_priority: Normal
---


# TextBox.AsianLineBreak property (Access)

Returns or sets a **Boolean** indicating whether line breaks in text boxes follow rules governing East Asian languages. **True** to control line breaks based on East Asian language rules. Read/write.


## Syntax

_expression_.**AsianLineBreak**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

Setting the **AsianLineBreak** property to **True** moves any punctuation marks and closing parentheses at the beginning of a line to the end of the previous line, and moves opening parentheses at the end of a line to the beginning of the next line.


## Example

This example sets all the text boxes on the specified form to break lines according to East Asian language rules.

```vb
Dim ctlLoop As Control 
 
For Each ctlLoop In Forms(0).Controls 
 If ctlLoop.ControlType = acTextBox Then 
 ctlLoop.AsianLineBreak = True 
 End If 
Next ctlLoop
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]