---
title: StyleSheet.FullName property (Word)
keywords: vbawd10.chm166658049
f1_keywords:
- vbawd10.chm166658049
ms.prod: word
api_name:
- Word.StyleSheet.FullName
ms.assetid: 81b79219-1aaf-c38b-4d78-62d7364f7374
ms.date: 06/08/2017
localization_priority: Normal
---


# StyleSheet.FullName property (Word)

Specifies the name of a cascading style sheet, including the drive or Web path. Read-only  **String**.


## Syntax

_expression_.**FullName**

_expression_ Required. A variable that represents a '[StyleSheet](Word.StyleSheet.md)' object.


## Remarks

Using this property is equivalent to using the **Path**, **PathSeparator**, and **Name** properties in sequence.


## Example

This example displays the path and file name of the style sheet attached to the active document.


```vb
Sub CSSName() 
 MsgBox ActiveDocument.StyleSheets(1).FullName 
End Sub
```


## See also


[StyleSheet Object](Word.StyleSheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]