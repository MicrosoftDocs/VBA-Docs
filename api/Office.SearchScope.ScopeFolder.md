---
title: SearchScope.ScopeFolder property (Office)
keywords: vbaof11.chm251002
f1_keywords:
- vbaof11.chm251002
ms.prod: office
api_name:
- Office.SearchScope.ScopeFolder
ms.assetid: 9bb05a24-7d9c-e218-40b1-06c054baacab
ms.date: 01/23/2019
localization_priority: Normal
---


# SearchScope.ScopeFolder property (Office)

Gets a **ScopeFolder** object. Read-only.


## Syntax

_expression_.**ScopeFolder**

_expression_ A variable that represents a **[SearchScope](Office.SearchScope.md)** object.


## Example

The following example displays the root path of each directory in My Computer. To retrieve this information, the example first gets the **ScopeFolder** object at the root of My Computer. The path of this **ScopeFolder** will always be "*". As with all **ScopeFolder** objects, the root object contains a **ScopeFolders** collection. This example loops through this **ScopeFolders** collection and displays the path of each **ScopeFolder** object in it. The paths of these **ScopeFolder** objects will be `A:\`, `C:\`, etc.


```vb
Sub DisplayRootScopeFolders() 
 
 'Declare variables that reference a 
 'SearchScope and a ScopeFolder object. 
 Dim ss As SearchScope 
 Dim sf As ScopeFolder 
 
 'Loop through the SearchScopes collection 
 'and display all of the root ScopeFolders collections in 
 'the My Computer scope. 
 For Each ss In SearchScopes 
 Select Case ss.Type 
 Case msoSearchInMyComputer 
 
 'Loop through each ScopeFolder object in 
 'the ScopeFolders collection of the 
 'SearchScope object and display the path. 
 For Each sf In ss.ScopeFolder.ScopeFolders 
 MsgBox "Path: " & sf.Path 
 Next sf 
 
 Case Else 
 End Select 
 Next ss 
 
End Sub
```


## See also

- [SearchScope object members](overview/Library-Reference/searchscope-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
