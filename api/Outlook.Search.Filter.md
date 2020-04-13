---
title: Search.Filter property (Outlook)
keywords: vbaol11.chm2253
f1_keywords:
- vbaol11.chm2253
ms.prod: outlook
api_name:
- Outlook.Search.Filter
ms.assetid: f6040465-da73-56f6-edb7-06d93bb8b531
ms.date: 06/08/2017
localization_priority: Normal
---


# Search.Filter property (Outlook)

Returns a **String** value that represents the DASL statement used to restrict the search to a specified subset of data. Read-only


## Syntax

_expression_. `Filter`

_expression_ A variable that represents a [Search](Outlook.Search.md) object.


## Remarks

This property is set as the  _Filter_ argument in the **[Application](Outlook.Application.md)** object's **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method.

When searching  **Text** fields, you can use either an apostrophe (') or double quotation marks ("") to delimit the values that are part of the filter. For example, all of the following lines function correctly when the field is of type **String** :




```vb
sFilter = "[CompanyName] = 'Microsoft'"
```




```vb
sFilter = "[CompanyName] = ""Microsoft"""
```




```vb
sFilter = "[CompanyName] = " & Chr(34) & "Microsoft" & Chr(34)
```


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new **[Search](Outlook.Search.md)** object. The event subroutine fires after the search has finished and displays the **[Tag](Outlook.Search.Tag.md)** and **Filter** properties of the **Search** object in addition to the results of the search.


```vb
Sub SearchInboxFolder() 
 
 'Searches the Inbox folder 
 
 Dim objSch As Outlook.Search 
 
 Const strF As String = _ 
 
 "urn:schemas:mailheader:subject = 'Office Holiday Party'" 
 
 Const strS As String = "Inbox" 
 
 Const strTag As String = "SubjectSearch" 
 
 Set objSch = _ 
 
 Application.AdvancedSearch(Scope:=strS, Filter:=strF, Tag:=strTag) 
 
End Sub
```

Use an **[AdvancedSearchComplete](Outlook.Application.AdvancedSearchComplete.md)** event subroutine to ensure the integrity of the data stored in the **Search** object.




```vb
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 Dim objRsts As Outlook.Results 
 
 Dim Item as Outlook.MailItem 
 
 MsgBox "The search " & SearchObject.Tag & "has finished. The filter used was " & _ 
 
 SearchObject.Filter & "." 
 
 Set objRsts = SearchObject.Results 
 
 'Print out number in results collection 
 
 MsgBox objRsts.Count 
 
 'Print out each member of results collection 
 
 For Each Item In objRsts 
 
 MsgBox Item 
 
 Next 
 
 
 
End Sub
```


## See also


[Search Object](Outlook.Search.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]