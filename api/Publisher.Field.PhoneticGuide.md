---
title: Field.PhoneticGuide property (Publisher)
keywords: vbapb10.chm6094856
f1_keywords:
- vbapb10.chm6094856
ms.prod: publisher
api_name:
- Publisher.Field.PhoneticGuide
ms.assetid: c68dba00-56c0-abba-0be8-5bd1a725f678
ms.date: 06/07/2019
localization_priority: Normal
---


# Field.PhoneticGuide property (Publisher)

Returns a **[PhoneticGuide](publisher.phoneticguide.md)** object that represents the properties of phonetic text applied to a text range.


## Syntax

_expression_.**PhoneticGuide**

_expression_ A variable that represents a **[Field](Publisher.Field.md)** object.


## Return value

PhoneticGuide


## Example

This example adds phonetic text to the selection and displays the text to which the phonetic text applies, which is the originally selected text. This example assumes that text is selected. If no text is selected, the message box is blank.

```vb
Sub AddPhoneticText() 
 With Selection.TextRange.Fields.AddPhoneticGuide _ 
 (Range:=Selection.TextRange, Text:="ver-E nIs") 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]