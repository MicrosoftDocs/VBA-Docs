---
title: Application.WordBasic property (Word)
keywords: vbawd10.chm158334977
f1_keywords:
- vbawd10.chm158334977
ms.prod: word
api_name:
- Word.Application.WordBasic
ms.assetid: 8c405ea6-0073-f994-42b2-cacb986f1f1f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WordBasic property (Word)

Returns an automation object (WordBasic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.


## Syntax

_expression_. `WordBasic`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

In Word 2000 and later, when you open a Word version 6.0 or Word for Windows 95 template that contains WordBasic macros, the macros are automatically converted to Visual Basic modules. Each WordBasic statement and function in the macro is converted to the corresponding WordBasic method.

For information about WordBasic statements and functions, see the WordBasic Help in Word version 6.0 or Word for Windows 95. For information about converting WordBasic to Visual Basic, see [Converting WordBasic Macros to Visual Basic](../word/Concepts/Customizing-Word/converting-wordbasic-macros-to-visual-basic.md). For general information, see [Conceptual Differences Between WordBasic and Visual Basic](../word/Concepts/Customizing-Word/conceptual-differences-between-wordbasic-and-visual-basic.md).


## Example

This example uses the WordBasic object to create a new document in Word version 6.0 or Word for Windows 95 and insert the available font names. Each font name is formatted in its corresponding font.


```vb
With WordBasic 
 .FileNewDefault 
 For aCount = 1 To .CountFonts() 
 .Font .[Font$](aCount) 
 .Insert .[Font$](aCount) 
 .InsertPara 
 Next 
End With
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]