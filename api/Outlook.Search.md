---
title: Search object (Outlook)
keywords: vbaol11.chm2248
f1_keywords:
- vbaol11.chm2248
ms.prod: outlook
api_name:
- Outlook.Search
ms.assetid: 226a5d49-3caf-90dd-725c-265404d1939f
ms.date: 06/08/2017
localization_priority: Normal
---


# Search object (Outlook)

Contains information about individual searches performed against Outlook items.


## Remarks

The  **Search** object contains properties that define the type of search and the parameters of the search itself.

Use the  **[Application](Outlook.Application.md)** object's **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method to return a **Search** object.

Use the  **[AdvancedSearchComplete](Outlook.Application.AdvancedSearchComplete.md)** event to determine when a given search has completed.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns a search object named "SubjectSearch" and displays the object's  **[Tag](Outlook.Search.Tag.md)** and **[Filter](Outlook.Search.Filter.md)** property values. The **Tag** property is used to identify a specific search once it has completed.


```vb
Sub SearchInboxFolder() 
 
'Searches the Inbox 
 
 
 
 Dim objSch As Search 
 
 Const strF As String = _ 
 
 "urn:schemas:mailheader:subject = 'Office Christmas Party'" 
 
 Const strS As String = "Inbox" 
 
 Const strTag As String = "SubjectSearch" 
 
 Set objSch = Application.AdvancedSearch(Scope:=strS, _ 
 
 Filter:=strF, SearchSubFolders:=True, Tag:=strTag) 
 
 
 
End Sub 
 

```

The following VBA example displays information about the search and the results of the search.




```vb
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 
 
 Dim objRsts As Results 
 
 MsgBox "The search " & SearchObject.Tag & "has completed. 
 
 Set objRsts = SearchObject.Results 
 
 'Print out number in Results collection 
 
 Debug.Print objRsts.Count 
 
 'Print out each member of Results collection 
 
 For Each Item In objRsts 
 
 Debug.Print Item 
 
 Next 
 
 
 
End Sub 
 

```


## Methods



|Name|
|:-----|
|[GetTable](Outlook.Search.GetTable.md)|
|[Save](Outlook.Search.Save.md)|
|[Stop](Outlook.Search.Stop.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Search.Application.md)|
|[Class](Outlook.Search.Class.md)|
|[Filter](Outlook.Search.Filter.md)|
|[IsSynchronous](Outlook.Search.IsSynchronous.md)|
|[Parent](Outlook.Search.Parent.md)|
|[Results](Outlook.Search.Results.md)|
|[Scope](Outlook.Search.Scope.md)|
|[SearchSubFolders](Outlook.Search.SearchSubFolders.md)|
|[Session](Outlook.Search.Session.md)|
|[Tag](Outlook.Search.Tag.md)|

## See also


[Search Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]