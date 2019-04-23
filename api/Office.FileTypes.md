---
title: FileTypes object (Office)
keywords: vbaof11.chm257000
f1_keywords:
- vbaof11.chm257000
ms.prod: office
api_name:
- Office.FileTypes
ms.assetid: 5e8b5240-5ebd-704d-72e6-1f4ad951dfdc
ms.date: 01/09/2019
localization_priority: Normal
---


# FileTypes object (Office)

A collection of values of the type **msoFileType** that determine which types of files are returned during a search.


## Remarks

There is only one **FileTypes** collection for all searches, so it's important to clear the **FileTypes** collection before executing a search unless you wish to search for file types from previous searches. 

The easiest way to clear the collection is to set the **FileType** property to the first file type for which you want to search. You can also remove individual types by using the **Remove** method. To determine the file type of each item in the collection, use the **Item** property to return the **msoFileType** value.


## See also

- [FileTypes object members](overview/library-reference/filetypes-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]