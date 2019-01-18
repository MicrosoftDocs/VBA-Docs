---
title: Envelope.DefaultOrientation property (Word)
keywords: vbawd10.chm152567827
f1_keywords:
- vbawd10.chm152567827
ms.prod: word
api_name:
- Word.Envelope.DefaultOrientation
ms.assetid: b227ba33-0114-db43-9d5e-a18e6b8a868a
ms.date: 06/08/2017
localization_priority: Normal
---


# Envelope.DefaultOrientation property (Word)

Returns or sets the default orientation for feeding envelopes. Read/write  **WdEnvelopeOrientation**.


## Syntax

 _expression_. `DefaultOrientation`

 _expression_ Required. A variable that represents an '[Envelope](Word.Envelope.md)' object.


## Example

This example sets envelopes to be fed face up, centered, and in portrait orientation.


```vb
With ActiveDocument.Envelope 
 .DefaultFaceUp = True 
 .DefaultOrientation = wdCenterPortrait 
 MsgBox "Feed envelopes centered, face up, " _ 
 & "in portrait orientation" 
End With
```


## See also


[Envelope Object](Word.Envelope.md)

