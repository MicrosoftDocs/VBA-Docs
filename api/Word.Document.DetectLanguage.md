---
title: Document.DetectLanguage method (Word)
keywords: vbawd10.chm158007447
f1_keywords:
- vbawd10.chm158007447
ms.prod: word
api_name:
- Word.Document.DetectLanguage
ms.assetid: 625cff5b-630e-bcaa-1094-57db5029ebd9
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.DetectLanguage method (Word)

Analyzes the specified text to determine the language that it is written in.


## Syntax

_expression_. `DetectLanguage`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

When applied to a  **Document** object, the **DetectLanguage** method checks all available text in the document (headers, footers, text boxes, and so forth). If the specified text contains a partial sentence, the selection or range is extended to the end of the sentence.

If the  **DetectLanguage** method has already been applied to the specified text, the **[LanguageDetected](Word.Document.LanguageDetected.md)** property is set to **True**. To re-evaluate the language of the specified text, you must first set the **LanguageDetected** property to **False**.


## Example

This example checks the active document to determine the language it's written in and then displays the result.


```vb
With ActiveDocument 
 If .LanguageDetected = True Then 
 x = MsgBox("This document has already " _ 
 & "been checked. Do you want to check " _ 
 & "it again?", vbYesNo) 
 If x = vbYes Then 
 .LanguageDetected = False 
 .DetectLanguage 
 End If 
 Else 
 .DetectLanguage 
 End If 
 If .Range.LanguageID = wdEnglishUS Then 
 MsgBox "This is a U.S. English document." 
 Else 
 MsgBox "This is not a U.S. English document." 
 End If 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]