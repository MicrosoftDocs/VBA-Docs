---
title: Application.IsSearchSynchronous method (Outlook)
keywords: vbaol11.chm729
f1_keywords:
- vbaol11.chm729
ms.prod: outlook
api_name:
- Outlook.Application.IsSearchSynchronous
ms.assetid: cd757b43-5e3f-1504-9944-7431bda6f004
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsSearchSynchronous method (Outlook)

Returns a **Boolean** indicating if a search will be synchronous or asynchronous.


## Syntax

_expression_. `IsSearchSynchronous`( `_LookInFolders_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _LookInFolders_|Required| **String**|The path name of the folders that the search will search through. You must enclose the folder path with single quotes.|

## Return value

 **True** if the search is synchronous; otherwise, **False**.


## Remarks

If the search is synchronous, the  **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method will not return until the search has completed. Conversely, if the search is asynchronous, the **AdvancedSearch** method will immediately return. In order to get meaningful results from an asynchronous search, use the **[AdvancedSearchComplete](Outlook.Application.AdvancedSearchComplete.md)** event to notify you when the search has finished.


## Example




```vb
Sub TestStoresForSynchronousSearch() 
 
 Dim folderPath As String 
 
 Dim oStore As Outlook.Store 
 
 For Each oStore In Outlook.Session.Stores 
 
 folderPath = "'" & oStore.GetRootFolder.folderPath & "'" 
 
 Debug.Print folderPath & " IsSearchSynchronous = " & _ 
 
 Application.IsSearchSynchronous(folderPath) 
 
 Next 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]