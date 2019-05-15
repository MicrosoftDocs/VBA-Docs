---
title: Worksheet.CommentsThreaded property (Excel)
ms.prod: excel
api_name:
- Excel.Worksheet.CommentsThreaded
ms.date: 05/15/2019
localization_priority: Normal
---


# Worksheet.CommentsThreaded property (Excel)

Returns a **[CommentsThreaded](Excel.CommentsThreaded.md)** collection that represents all the top-level/root comments (no replies) for the specified worksheet. Includes legacy and modern comments. Read-only. 

## Syntax

_expression_.**CommentsThreaded**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example deletes all threaded comments added by author Jean Selva on the active sheet.

```vb
For Each c in ActiveSheet.CommentsThreaded
 If c.Author.Name = "Jean Selva" Then c.Delete 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]