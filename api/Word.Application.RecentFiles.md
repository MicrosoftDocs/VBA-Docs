---
title: Application.RecentFiles property (Word)
keywords: vbawd10.chm158334983
f1_keywords:
- vbawd10.chm158334983
ms.prod: word
api_name:
- Word.Application.RecentFiles
ms.assetid: 517fb0cf-2dfb-f0a0-0882-f233198768d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RecentFiles property (Word)

Returns a  **[RecentFiles](Word.recentfiles.md)** collection that represents the most recently accessed files.


## Syntax

_expression_.**RecentFiles**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example opens the first item in the **RecentFiles** collection (the first document name listed on the **File** menu).


```vb
If RecentFiles.Count >= 1 Then RecentFiles(1).Open
```

This example displays the name of each file in the **RecentFiles** collection.




```vb
For Each rFile In RecentFiles 
 MsgBox rFile.Name 
Next rFile
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]