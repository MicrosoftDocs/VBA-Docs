---
title: Range.CopyAsPicture method (Word)
keywords: vbawd10.chm157155495
f1_keywords:
- vbawd10.chm157155495
ms.prod: word
api_name:
- Word.Range.CopyAsPicture
ms.assetid: b104bb78-9e76-37c7-2102-f71a3d8ddabb
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.CopyAsPicture method (Word)

The **CopyAsPicture** method works the same way as the **Copy** method.


## Syntax

_expression_. `CopyAsPicture`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example copies the contents of the active document as a picture and pastes it as a picture at the end of the document.


```vb
Sub CopyPasteAsPicture() 
 With ActiveDocument.Range 
 .CopyAsPicture 
 .Collapse Direction:=wdCollapseEnd 
 .PasteSpecial DataType:=wdPasteMetafilePicture 
 End With 
End Sub
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]