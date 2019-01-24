---
title: SearchFolders object (Office)
keywords: vbaof11.chm258000
f1_keywords:
- vbaof11.chm258000
ms.prod: office
api_name:
- Office.SearchFolders
ms.assetid: 5958cafc-880e-ee9f-b2f5-be463bfe5232
ms.date: 01/23/2019
localization_priority: Normal
---


# SearchFolders object (Office)

A collection of **[ScopeFolder](Office.ScopeFolder.md)** objects that determines which folders are searched.


## Remarks

For each application, there is only a single **SearchFolders** collection. The contents of the collection remains after the code that calls it has finished executing. Consequently, it is important to clear the collection unless you want to include folders from previous searches in your search.

You can use the **Add** method of the **SearchFolders** collection to add a **ScopeFolder** object to the **SearchFolders** collection; however, it is usually simpler to use the **[AddToSearchFolders](office.scopefolder.addtosearchfolders.md)** method of the **ScopeFolder** that you want to add because there is only one **SearchFolders** collection for all searches.


## See also

- [SearchFolders object members](overview/Library-Reference/searchfolders-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]