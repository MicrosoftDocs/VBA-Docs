---
title: Envelope object (Word)
keywords: vbawd10.chm2328
f1_keywords:
- vbawd10.chm2328
ms.prod: word
api_name:
- Word.Envelope
ms.assetid: 03664453-f7fb-f76a-ea60-37e72b53e17c
ms.date: 06/08/2017
localization_priority: Normal
---


# Envelope object (Word)

Represents an envelope attached to a document.


## Remarks

Use the **[Envelope](Word.Document.Envelope.md)** property to return the **Envelope** object. The following example adds an envelope to a new document and sets the distance between the top of the envelope and the address to 2.25 inches.


```vb
Set myDoc = Documents.Add 
addr = "Michael Matey" & vbCr & "123 Skye St." _ 
 & vbCr & "Redmond, WA 98107" 
retaddr = "Cora Edmonds" & vbCr & "456 Erde Lane" & vbCr _ 
 & "Redmond, WA 98107" 
With myDoc.Envelope 
 .Insert Address:=addr, ReturnAddress:=retaddr 
 .AddressFromTop = InchesToPoints(2.25) 
End With
```

Remarks

The **Envelope** object is available regardless of whether an envelope has been added to the specified document. However, an error occurs if you use one of the following properties when an envelope has not been added to the document: **[Address](Word.Envelope.Address.md)**, **[AddressFromLeft](Word.Envelope.AddressFromLeft.md)**, **[AddressFromTop](Word.Envelope.AddressFromTop.md)**, **[FeedSource](Word.Envelope.FeedSource.md)**, **[ReturnAddress](Word.Envelope.ReturnAddress.md)**, **[ReturnAddressFromLeft](Word.Envelope.ReturnAddressFromLeft.md)**, **[ReturnAddressFromTop](Word.Envelope.ReturnAddressFromTop.md)**, and **[UpdateDocument](Word.Envelope.UpdateDocument.md)**.

The following example demonstrates how to use the **On Error GoTo** statement to trap the error that occurs if an envelope has not been added to the active document. If, however, an envelope has been added to the document, the recipient address is displayed.




```vb
On Error GoTo ErrorHandler 
MsgBox ActiveDocument.Envelope.Address 
ErrorHandler: 
If Err = 5852 Then MsgBox _ 
 "Envelope is not in the specified document"
```

Use the **Insert** method to add an envelope to the specified document. Use the **PrintOut** method to set the properties of an envelope and print it without adding it to the document.


> [!NOTE] 
> There is no Envelopes collection; each  **Document** object contains only one **Envelope** object.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]