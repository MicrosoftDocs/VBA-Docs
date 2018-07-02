---
title: ScopeFolder Object (Office)
keywords: vbaof11.chm259000
f1_keywords:
- vbaof11.chm259000
ms.prod: office
api_name:
- Office.ScopeFolder
ms.assetid: fe46c1ad-fd60-a698-23dd-04d0631ac403
ms.date: 06/08/2017
---


# ScopeFolder Object (Office)

Corresponds to a searchable folder.  **ScopeFolder** objects are intended for use with the **SearchFolders** collection.


## Remarks

When you want to search specific folders you can use the methods and properties of the  **SearchScope** object and **ScopeFolders** collection to retrieve **ScopeFolder** objects and add them to the **SearchFolders** collection.

In each  **ScopeFolder** object there is a **ScopeFolders** collection that contains the subfolders of the parent **ScopeFolder** object. You can traverse the entire folder structure of a search scope (for example, all local drives) by looping through these **ScopeFolders** collections and returning all of the lower-level **ScopeFolder** objects. A **ScopeFolder** object with no subfolders contains an empty **ScopeFolders** collection.

For an example that demonstrates how to loop through all of the  **ScopeFolder** objects in a search scope, see the **SearchFolders** collection topic.

You can use the  **Add** method of the **SearchFolders** collection to add a **ScopeFolder** object to the **SearchFolders** collection, however, it is usually simpler to use the **AddToSearchFolders** method of the **ScopeFolder** that you want to add, as there is only one **SearchFolders** collection for all searches.

For an example that demonstrates how to add a  **ScopeFolder** to the **SearchFolders** collection, see the **SearchFolders** collection topic.


## Example

Use the  **ScopeFolder** property of the **SearchScope** object to return the root **ScopeFolder** object of a search scope; for example:


```vb
Set sf = SearchScopes.Item(1).ScopeFolder
```

Use the  **Item** method of the **ScopeFolders** collection to return a subfolder of a root **ScopeFolder** object; for example:




```vb
Set sf = SearchScopes.Item(1).ScopeFolder.ScopeFolders.Item(1)
```

The following example displays the root path of each directory in My Computer. To retrieve this information, the example first gets the  **ScopeFolder** object at the root of My Computer. The path of this **ScopeFolder** object will always be "*". As with all **ScopeFolder** objects, the root object contains a **ScopeFolders** collection. This example loops through this **ScopeFolders** collection and displays the path of each **ScopeFolder** object in it. The paths of these **ScopeFolder** objects will be "A:\", "C:\", etc.




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
 MsgBox "ScopeFolder object's path: " &amp; sf.Path 
 Next sf 
 
 Case Else 
 End Select 
 Next 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddToSearchFolders](Office.ScopeFolder.AddToSearchFolders.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.ScopeFolder.Application.md)|
|[Creator](Office.ScopeFolder.Creator.md)|
|[Name](Office.ScopeFolder.Name.md)|
|[Path](Office.ScopeFolder.Path.md)|
|[ScopeFolders](Office.ScopeFolder.ScopeFolders.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
