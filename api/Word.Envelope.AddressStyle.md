---
title: Envelope.AddressStyle property (Word)
keywords: vbawd10.chm152567825
f1_keywords:
- vbawd10.chm152567825
ms.prod: word
api_name:
- Word.Envelope.AddressStyle
ms.assetid: 404962d4-18eb-f79a-67e4-e54c3d6539e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Envelope.AddressStyle property (Word)

Returns a  **[Style](Word.Style.md)** object that represents the delivery address style for the envelope. Read-only.


## Syntax

 _expression_. `AddressStyle`

 _expression_ A variable that represents a '[Envelope](Word.Envelope.md)' object.


## Remarks

If an envelope is added to the document, text formatted with the Envelope Address style is automatically updated.


## Example

This example modifies the font formatting associated with the Envelope Address style.


```vb
With ActiveDocument.Envelope.AddressStyle.Font 
 .Bold = False 
 .Name = "Times New Roman" 
 .Size = 16 
End With
```


## See also


[Envelope Object](Word.Envelope.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]