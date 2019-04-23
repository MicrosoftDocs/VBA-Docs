---
title: Document.SaveFormat property (Word)
keywords: vbawd10.chm158007355
f1_keywords:
- vbawd10.chm158007355
ms.prod: word
api_name:
- Word.Document.SaveFormat
ms.assetid: f8d31365-1935-307f-3663-d6e769944489
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SaveFormat property (Word)

Returns the file format of the specified document or file converter. Read-only  **Long**.


## Syntax

_expression_. `SaveFormat`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

The  **SaveFormat** property will be a unique number that specifies an external file converter or a **WdSaveFormat** constant.

Use the value of the  **SaveFormat** property for the _FileFormat_ argument of the **[SaveAs2](Word.SaveAs2.md)** method to save a document in a file format for which there isn't a corresponding **WdSaveFormat** constant.


## Example

If the active document is a Rich Text Format (RTF) document, this example saves it as a Microsoft Word document.


```vb
If ActiveDocument.SaveFormat = wdFormatRTF Then 
 ActiveDocument.SaveAs FileFormat:=wdFormatDocument 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]