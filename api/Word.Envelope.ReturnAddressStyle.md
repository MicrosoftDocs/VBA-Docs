---
title: Envelope.ReturnAddressStyle property (Word)
keywords: vbawd10.chm152567826
f1_keywords:
- vbawd10.chm152567826
ms.prod: word
api_name:
- Word.Envelope.ReturnAddressStyle
ms.assetid: cebc53db-5c79-c036-7e15-835095affbde
ms.date: 06/08/2017
localization_priority: Normal
---


# Envelope.ReturnAddressStyle property (Word)

Returns a  **[Style](Word.Style.md)** object that represents the return address style for the envelope.


## Syntax

 _expression_. `ReturnAddressStyle`

 _expression_ An expression that returns an '[Envelope](Word.Envelope.md)' object.


## Remarks

If an envelope is added to the document, text formatted with the Envelope Return style is automatically updated.


## Example

This example displays the style name and description of the envelope return address.


```vb
Set myStyle = ActiveDocument.Envelope.ReturnAddressStyle 
MsgBox Prompt:=myStyle.Description, Title:=myStyle.NameLocal
```

This example sets the line spacing and space-after formatting for the envelope return address style.




```vb
With ActiveDocument.Envelope.ReturnAddressStyle.ParagraphFormat 
 .LineSpacingRule = wdLineSpaceExactly 
 .LineSpacing = 13 
 .SpaceAfter = 6 
End With
```


## See also


[Envelope Object](Word.Envelope.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]