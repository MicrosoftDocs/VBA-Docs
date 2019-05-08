---
title: MailMergeDataSource.HeaderSourceName property (Word)
keywords: vbawd10.chm152895490
f1_keywords:
- vbawd10.chm152895490
ms.prod: word
api_name:
- Word.MailMergeDataSource.HeaderSourceName
ms.assetid: 80380230-3f88-f08d-780b-923287d359fa
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeDataSource.HeaderSourceName property (Word)

Returns the path and file name of the header source attached to the specified mail merge main document. Read-only  **String**.


## Syntax

_expression_. `HeaderSourceName`

_expression_ A variable that represents a '[MailMergeDataSource](Word.MailMergeDataSource.md)' object.


## Example

If a header source is attached to the active document, this example displays the file name.


```vb
Dim strName As String 
 
strName = ActiveDocument.MailMerge.DataSource.HeaderSourceName 
If strName <> "" Then MsgBox strName
```

This example opens the header source attached to the active document if the source is a Word document.




```vb
Dim mmdsTemp As MailMergeDataSource 
 
Set mmdsTemp = ActiveDocument.MailMerge.DataSource 
 
If mmdsTemp.HeaderSourceType = wdMergeInfoFromWord Then 
 Documents.Open FileName:=mmdsTemp.HeaderSourceName 
End If
```


## See also


[MailMergeDataSource Object](Word.MailMergeDataSource.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]