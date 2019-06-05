---
title: Document.PrintOutEx method (Publisher)
keywords: vbapb10.chm196755
f1_keywords:
- vbapb10.chm196755
ms.prod: publisher
api_name:
- Publisher.Document.PrintOutEx
ms.assetid: f11b6f8b-08a0-28f6-5930-47d684585bef
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.PrintOutEx method (Publisher)

Prints all or part of the specified publication.


## Syntax

_expression_.**PrintOut** (_From_, _To_, _PrintToFile_, _Copies_, _Collate_, _PrintStyle_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_From_|Optional| **Long**|The starting page number.|
|_To_|Optional| **Long**|The ending page number.|
|_PrintToFile_|Optional| **String**|The path and file name of a document to be printed to a file.|
|_Copies_|Optional| **Long**|The number of copies to be printed.|
|_Collate_|Optional| **Boolean**|When printing multiple copies of a document, **True** to print all pages of the document before printing the next copy.|
|_PrintStyle_|Optional| **[PbPrintStyle](Publisher.PbPrintStyle.md)**|The print style to use. Can be one of the **PbPrintStyle** constants declared in the Microsoft Publisher type library.|

## Remarks

If _PrintStyle_ is **pbPrintStyleMultipleCopiesPerSheet** or **pbPrintStyleMultiplePagesPerSheet**, Publisher ignores any value that you pass for the _Collate_ parameter.


## Example

This example prints the active publication.

```vb
Sub PrintActivePublication() 
 ThisDocument.PrintOutEx 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]