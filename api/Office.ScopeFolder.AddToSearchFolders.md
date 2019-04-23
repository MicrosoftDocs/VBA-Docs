---
title: ScopeFolder.AddToSearchFolders method (Office)
keywords: vbaof11.chm259004
f1_keywords:
- vbaof11.chm259004
ms.prod: office
api_name:
- Office.ScopeFolder.AddToSearchFolders
ms.assetid: e77e2406-b709-0f3e-736d-2fd56c7447e1
ms.date: 01/23/2019
localization_priority: Normal
---


# ScopeFolder.AddToSearchFolders method (Office)

Adds a **ScopeFolder** object to the **[SearchFolders](office.searchfolders.md)** collection.


## Syntax

_expression_.**AddToSearchFolders**

_expression_ A variable that represents a **[ScopeFolder](Office.ScopeFolder.md)** object.


## Remarks

Although you can use the **SearchFolders** collection's **Add** method to add a **ScopeFolder** object to the **SearchFolders** collection, it is usually simpler to use the **AddToSearchFolders** method of the **ScopeFolder** object that you want to add, because there is only one **SearchFolders** collection for all searches.


## Example

The following example adds the root **ScopeFolder** object to the **SearchFolders** collection. For a longer example that uses the **AddToSearchFolders** method, see the **SearchFolders** collection topic.


```vb
SearchScopes(1).ScopeFolder.AddToSearchFolders
```


## See also

- [ScopeFolder object members](overview/Library-Reference/scopefolder-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
