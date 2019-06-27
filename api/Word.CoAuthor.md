---
title: CoAuthor object (Word)
keywords: vbawd10.chm1237
f1_keywords:
- vbawd10.chm1237
ms.prod: word
api_name:
- Word.CoAuthor
ms.assetid: d1b58eea-4570-ffd3-4c13-a74a998b079e
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthor object (Word)

Represents a single co author in the document. The  **CoAuthor** object is a member of the **[CoAuthors](Word.CoAuthors.md)** collection. The **CoAuthors** collection contains all the co authors in the document (authors that are actively editing the document).


> [!IMPORTANT] 
> Documents can only be co authored on a server that supports the File Synchronization via SOAP over HTTP protocol, such as Microsoft SharePoint Server 2010.


## Remarks

Use  **CoAuthors** (_index_), where _index_ is the index number to return a single **CoAuthor** object.


> [!NOTE] 
> When a new co author begins to edit the document, it can take up to one minute or longer for the co author to appear in the document.


## Example

The following code example returns the name of the first co author in the active document.


```vb
Dim author As CoAuthor 
 
Set author = ActiveDocument.CoAuthoring.Authors(1) 
MsgBox "The name of the first co author in this document is " & author.Name
```


## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]