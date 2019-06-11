---
title: Page.IsLeading property (Publisher)
keywords: vbapb10.chm131102
f1_keywords:
- vbapb10.chm131102
ms.prod: publisher
api_name:
- Publisher.Page.IsLeading
ms.assetid: 5a65f1fe-442d-f352-bea6-b732771008d8
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.IsLeading property (Publisher)

**True** if the specified **Page** object is a leading page of a two-page spread. Read-only **Boolean**.


## Syntax

_expression_.**IsLeading**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

Boolean


## Example

The following example displays for each page whether the page is a trailing or leading page in the publication.

```vb
Dim objPage As Page 
Dim strPageInfo As String 
For Each objPage In ActiveDocument.Pages 
 strPageInfo = "Page number " & objPage.PageNumber 
 If objPage.IsLeading Then 
 strPageInfo = strPageInfo & " is a leading page." & Chr(13) 
 ElseIf objPage.IsTrailing Then 
 strPageInfo = strPageInfo & " is a trailing page." & Chr(13) 
 End If 
 MsgBox strPageInfo 
Next objPage
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]