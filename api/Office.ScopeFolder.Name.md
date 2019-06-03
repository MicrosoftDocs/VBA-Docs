---
title: ScopeFolder.Name property (Office)
keywords: vbaof11.chm259001
f1_keywords:
- vbaof11.chm259001
ms.prod: office
api_name:
- Office.ScopeFolder.Name
ms.assetid: da1cc239-2988-2b57-11d1-8313ae3d5566
ms.date: 01/23/2019
localization_priority: Normal
---


# ScopeFolder.Name property (Office)

Gets the name of a searchable folder. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[ScopeFolder](Office.ScopeFolder.md)** object.


## Return value

String


## Remarks

**ScopeFolder** objects are intended for use with the **[SearchFolders](office.searchfolders.md)** collection. The **SearchFolders** collection defines the folders that are searched.


## Example

The following example displays a message box with the name of the folder that will be searched.


```vb
Dim sf As ScopeFolder 
 Dim strScopeFolder As String 
 
 Set sf = SearchScopes.Item(1).ScopeFolder 
 strScopeFolder = sf.Name 
 
 MsgBox ("The name of the folder that will be searched is " & strScopeFolder) 

```


## See also

- [ScopeFolder object members](overview/Library-Reference/scopefolder-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
