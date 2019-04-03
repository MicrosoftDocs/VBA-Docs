---
title: Search.Scope property (Outlook)
keywords: vbaol11.chm2259
f1_keywords:
- vbaol11.chm2259
ms.prod: outlook
api_name:
- Outlook.Search.Scope
ms.assetid: aa4b9aea-029f-6f80-87b1-b99c04ff9631
ms.date: 06/08/2017
localization_priority: Normal
---


# Search.Scope property (Outlook)

Returns a  **String** that specifies the scope of the specified search. Read-only.


## Syntax

_expression_. `Scope`

_expression_ A variable that represents a [Search](Outlook.Search.md) object.


## Remarks

The scope of the search is defined when the search is initiated. For more information, see the  **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a  **Search** object. The user's **Inbox** is specified as the scope of the search. The event subroutine occurs when the search has completed and displays the **[Tag](Outlook.Search.Tag.md)** and **Scope** properties for the new object in addition to the results of the search.


```vb
Public blnSearchComp As Boolean 
 
 
 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 MsgBox "The AdvancedSearchComplete Event fired for " & SearchObject.Tag & " and the scope was " & SearchObject.Scope 
 
 blnSearchComp = True 
 
End Sub 
 
 
 
Sub TestAdvancedSearchComplete() 
 
 'List all items in the Inbox that do NOT have a flag. 
 
 Dim objSch As Outlook.Search 
 
 Const strF As String = "urn:schemas:httpmail:messageflag IS NULL" 
 
 Const strS As String = "Inbox" 
 
 Dim rsts As Outlook.Results 
 
 Dim i As Integer 
 
 blnSearchComp = False 
 
 Const strF1 As String = "urn:schemas:mailheader:subject = 'Test'" 
 
 Const strS1 As String = "Inbox" 
 
 Set objSch = _ 
 
 Application.AdvancedSearch(Scope:=strS1, Filter:=strF1, Tag:="FlagSearch") 
 
 While blnSearchComp = False 
 
 DoEvents 
 
 Wend 
 
 Set rsts = objSch.Results 
 
 For i = 1 To rsts.Count 
 
 MsgBox rsts.Item(i).SenderName 
 
 Next 
 
End Sub
```


## See also


[Search Object](Outlook.Search.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]