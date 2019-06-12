---
title: ReaderSpread.PageCount property (Publisher)
keywords: vbapb10.chm524294
f1_keywords:
- vbapb10.chm524294
ms.prod: publisher
api_name:
- Publisher.ReaderSpread.PageCount
ms.assetid: 39d26cd7-f4b8-bbf3-a2a8-32a4c9362e30
ms.date: 06/13/2019
localization_priority: Normal
---


# ReaderSpread.PageCount property (Publisher)

Returns a **Long** indicating the number of pages in the specified reader spread. Read-only.


## Syntax

_expression_.**PageCount**

_expression_ A variable that represents a **[ReaderSpread](Publisher.ReaderSpread.md)** object.


## Return value

Long


## Remarks

A reader spread can contain no more than two pages.


## Example

The following example checks the reader spread of the third page in the active publication to see if it contains more than one page, and then displays the total number of pages in the spread.

```vb
Sub NumberOfPagesInSpread() 
 If ActiveDocument.Pages(3).ReaderSpread.PageCount > 1 Then 
 MsgBox "The spread has two pages." 
 Else 
 MsgBox "The spread has only one page." 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]