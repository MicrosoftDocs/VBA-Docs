---
title: Selection.InsertFile method (Word)
keywords: vbawd10.chm158662779
f1_keywords:
- vbawd10.chm158662779
ms.prod: word
api_name:
- Word.Selection.InsertFile
ms.assetid: 963a5987-e6f8-824a-47d6-9788f026cf10
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertFile method (Word)

Inserts all or part of the specified file.


## Syntax

_expression_. `InsertFile`( `_FileName_` , `_Range_` , `_ConfirmConversions_` , `_Link_` , `_Attachment_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the file to be inserted. If you don't specify a path, Word assumes the file is in the current folder.|
| _Range_|Optional| **Variant**|If the specified file is a Word document, this parameter refers to a bookmark. If the file is another type (for example, a Microsoft Excel worksheet), this parameter refers to a named range or a cell range (for example, R1C1:R3C4).|
| _ConfirmConversions_|Optional| **Variant**| **True** to have Word prompt you to confirm conversion when inserting files in formats other than the Word Document format.|
| _Link_|Optional| **Variant**| **True** to insert the file by using an INCLUDETEXT field.|
| _Attachment_|Optional| **Variant**| **True** to insert the file as an attachment to an email message.|

## Example

This example uses an INCLUDETEXT field to insert the TEST.DOC file at the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseEnd 
Selection.InsertFile FileName:="C:\TEST.DOC", Link:=True
```

This example creates a new document and then inserts the contents of each text file in the C:\TMP folder into the new document.




```vb
Documents.Add 
ChDir "C:\TMP" 
myName = Dir("*.TXT") 
While myName <> "" 
 With Selection 
 .InsertFile FileName:=myName, ConfirmConversions:=False 
 .InsertParagraphAfter 
 .InsertBreak Type:=wdSectionBreakNextPage 
 .Collapse Direction:=wdCollapseEnd 
 End With 
 myName = Dir() 
Wend
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
