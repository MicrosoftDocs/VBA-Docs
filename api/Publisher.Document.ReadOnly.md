---
title: Document.ReadOnly property (Publisher)
keywords: vbapb10.chm196647
f1_keywords:
- vbapb10.chm196647
ms.prod: publisher
api_name:
- Publisher.Document.ReadOnly
ms.assetid: 9ee6488d-3070-e784-e772-78dace2c1284
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ReadOnly property (Publisher)

Returns **True** if the publication is read-only; returns **False** if it is read/write. Read-only **Boolean**.


## Syntax

_expression_.**ReadOnly**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

Boolean


## Example

This example saves the active publication and notifies the user that the file is saved and whether it is read-only.

```vb
Sub SaveAndStatus() 
 
 Dim bStatus As Boolean 
 
 Application.ActiveDocument.SaveAs "c:\testfile.pub" 
 bStatus = Application.ActiveDocument.ReadOnly 
 MsgBox "File Saved and Read-only Status = " & bStatus 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]