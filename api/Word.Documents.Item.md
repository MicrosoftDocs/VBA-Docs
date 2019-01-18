---
title: Documents.Item method (Word)
keywords: vbawd10.chm158072832
f1_keywords:
- vbawd10.chm158072832
ms.prod: word
api_name:
- Word.Documents.Item
ms.assetid: 0777c075-b466-3ac9-312a-4e1da7c1a732
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.Item method (Word)

Returns an individual  **Document** object in a collection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ Required. A variable that represents a '[Documents](Word.documents.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long** or **String**|The position within the collection of the object to be returned (**Long**), or the name of the object (**String**).|

## Return value

Document


## Example

This example displays the name of the first document in the  **Documents** collection.


```vb
Sub DocumentItem() 
 If Documents.Count >= 1 Then 
 MsgBox Documents.Item(1).Name 
 End If 
End Sub
```

This example saves the document named `Reports.doc`.

```vb
Documents.Item("Reports.docx").Save
```

## See also


[Documents Collection Object](Word.documents.md)

