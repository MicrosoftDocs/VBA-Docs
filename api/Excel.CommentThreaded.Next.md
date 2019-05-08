---
title: CommentThreaded.Next method (Excel)
keywords:
f1_keywords:
-
ms.prod: excel
api_name:
- Excel.CommentThreaded.Next
ms.assetid:
ms.date: 05/08/2019
localization_priority: Normal
---


# CommentThreaded.Next method (Excel)

Returns a **CommentThreaded** object that represents the next CommentThread.

## Syntax

_expression_.**Next**

_expression_ An expression that returns a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Return value

Comment Threaded


## Remarks

If called on a top-level (parent) **CommentThreaded object**, it will return a top-level (parent) **CommentThreaded** object that represents the next comment. Using this method on the last comment on a sheet returns Null (not the next comment on the next sheet).   

If called on a reply **CommentThreaded object**, it will return a reply **CommentThreaded** object that represents the next reply of a thread. This method works only on one thread. Using this method on the last reply of a thread returns a Null (not the next top-level comment). 


## Example

This example navigates to the next top level comment after the comment in range "A1" and updates its text.


```vb
Worksheets(1).Range("A1").CommentThreaded.Next.Text "CurrentSales"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]