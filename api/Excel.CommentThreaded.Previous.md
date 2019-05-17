---
title: CommentThreaded.Previous method (Excel)
keywords: vbaxl10.chm1010080
f1_keywords:
- vbaxl10.chm1010080
ms.prod: excel
api_name:
- Excel.CommentThreaded.Previous
ms.date: 05/15/2019
localization_priority: Normal
---


# CommentThreaded.Previous method (Excel)

Returns a **CommentThreaded** object that represents the previous threaded comment.

## Syntax

_expression_.**Previous**

_expression_ An expression that returns a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Return value

**CommentThreaded**


## Remarks

If called on a top-level (parent) **CommentThreaded** object, it returns a top-level (parent) **CommentThreaded** object that represents the previous comment. Using this method on the first comment on a sheet returns **Null** (not the last comment on the previous sheet).   

If called on a reply **CommentThreaded** object, it returns a reply **CommentThreaded** object that represents the previous reply of a thread. This method works only on one thread. Using this method on the first reply of a thread returns **Null** (not its top-level comment). 
 

## Example

This example navigates to the previous top-level comment after the comment in range E1, and updates its text.

```vb
Worksheets(1).Range("E1").CommentThreaded.Previous.Text "CurrentSales"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]