---
title: EmailSignatureEntries.Add Method (Word)
keywords: vbawd10.chm166002789
f1_keywords:
- vbawd10.chm166002789
ms.prod: word
api_name:
- Word.EmailSignatureEntries.Add
ms.assetid: da8b1a9a-aa3f-4288-887f-50d646d75728
ms.date: 06/08/2017
---


# EmailSignatureEntries.Add Method (Word)

Returns an  **[EmailSignatureEntry](Word.EmailSignatureEntry.md)** object that represents a new e-mail signature entry.


## Syntax

 _expression_. `Add`( `_Name_` , `_Range_` )

 _expression_ Required. A variable that represents an '[EmailSignatureEntries](Word.EmailSignatureEntries.md)' collection.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the e-mail entry.|
| _Range_|Required| **Range**|The range in the document that will be added as the signature.|

### Return value

EmailSignatureEntry


## Remarks

An e-mail signature is standard text that ends an e-mail message, such as your name and telephone number. Use the  **EmailSignatureEntries** property to create and manage a collection of e-mail signatures that Microsoft Word will use when creating e-mail messages.


## Example

This example adds an automatically numbered footnote at the end of the selection.


```vb
Sub NewSignature() 
 Application.EmailOptions.EmailSignature _ 
 .EmailSignatureEntries.Add _ 
 Name:=ActiveDocument.BuiltInDocumentProperties("Author"), _ 
 Range:=Selection.Range 
End Sub
```


## See also


[EmailSignatureEntries Collection](Word.EmailSignatureEntries.md)

