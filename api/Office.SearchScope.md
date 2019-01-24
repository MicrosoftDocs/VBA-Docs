---
title: SearchScope object (Office)
keywords: vbaof11.chm251000
f1_keywords:
- vbaof11.chm251000
ms.prod: office
api_name:
- Office.SearchScope
ms.assetid: 7faa5b49-6aa9-6682-165b-0d900fffd9ed
ms.date: 01/23/2019
localization_priority: Normal
---


# SearchScope object (Office)

Corresponds to a type of folder tree that can be searched.


## Remarks

Each **SearchScope** object contains a single **[ScopeFolder](office.scopefolder.md)** object that corresponds to the root folder of the search scope.

Use the **[Item](office.searchscopes.item.md)** property of the **SearchScopes** collection to return a **SearchScope** object; for example:

```vb
Dim ss As SearchScope 
Set ss = SearchScopes.Item(1)
```

Ultimately, the **SearchScope** object is intended to provide access to **ScopeFolder** objects that can be added to the **[SearchFolders](office.searchfolders.md)** collection. For an example that demonstrates how this is accomplished, see the **SearchFolders** collection topic.

See the **ScopeFolder** object topic to see a simple example of how to return a **ScopeFolder** object from a **SearchScope** object.


## Example

The following example displays all of the currently available **SearchScope** objects.


```vb
Sub DisplayAvailableScopes() 
 
 'Declare a variable that references a 
 'SearchScope object. 
 Dim ss As SearchScope 
 
 'Loop through the SearchScopes collection. 
 For Each ss In SearchScopes 
 Select Case ss.Type 
 Case msoSearchInMyComputer 
 MsgBox "My Computer is an available search scope." 
 Case msoSearchInMyNetworkPlaces 
 MsgBox "My Network Places is an available search scope." 
 Case msoSearchInOutlook 
 MsgBox "Outlook is an available search scope." 
 Case msoSearchInCustom 
 MsgBox "A custom search scope is available." 
 Case Else 
 MsgBox "Can't determine search scope." 
 End Select 
 Next ss 
 
End Sub
```


## See also

- [SearchScope object members](overview/Library-Reference/searchscope-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]