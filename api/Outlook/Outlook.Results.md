---
title: Results Object (Outlook)
keywords: vbaol11.chm3012
f1_keywords:
- vbaol11.chm3012
ms.prod: outlook
api_name:
- Outlook.Results
ms.assetid: 59057f6f-8f6d-eed0-c945-240b9593b7ea
ms.date: 06/08/2017
---


# Results Object (Outlook)

Contains data and results returned by the  **[Search](Outlook.Search.md)** object and the **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method.


## Remarks

The  **Results** object contains properties and methods that allow you to view and manipulate data. For example the **[GetNext](Outlook.Results.GetNext.md)**, **[GetPrevious](Outlook.Results.GetPrevious.md)**, **[GetFirst](Outlook.Results.GetFirst.md)**, and **[GetLast](Outlook.Results.GetLast.md)** methods allow you to search through the results and view the data by field. The **[Sort](Outlook.Results.Sort.md)** method allows you to sort the data.

Use the  **SearchObject.Results** property to return a **Results** object.


## Example

The following event procedure stores the results of a search in a variable named objRsts and displays the results of the search in the Immediate window.


```
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 
 
 Dim objRsts As Outlook.Results 
 
 MsgBox "The search " &amp; SearchObject.Tag &amp; _ 
 
 "has completed. The scope of the search was " &amp; _ 
 
 SearchObject.Scope &amp; "." 
 
 Set objRsts = SearchObject.Results 
 
 'Print out number in Results collection 
 
 Debug.Print objRsts.Count 
 
 'Print out each member of Results collection 
 
 For Each Item In objRsts 
 
 Debug.Print Item 
 
 Next 
 
 
 
End Sub 
 

```


## Events



|**Name**|
|:-----|
|[ItemAdd](Outlook.Results.ItemAdd.md)|
|[ItemChange](Outlook.Results.ItemChange.md)|
|[ItemRemove](Outlook.Results.ItemRemove.md)|

## Methods



|**Name**|
|:-----|
|[GetFirst](Outlook.Results.GetFirst.md)|
|[GetLast](Outlook.Results.GetLast.md)|
|[GetNext](Outlook.Results.GetNext.md)|
|[GetPrevious](Outlook.Results.GetPrevious.md)|
|[Item](Outlook.Results.Item.md)|
|[ResetColumns](Outlook.Results.ResetColumns.md)|
|[SetColumns](Outlook.Results.SetColumns.md)|
|[Sort](Outlook.Results.Sort.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.Results.Application.md)|
|[Class](results-class-property-outlook.md)|
|[Count](results-count-property-outlook.md)|
|[DefaultItemType](results-defaultitemtype-property-outlook.md)|
|[Parent](results-parent-property-outlook.md)|
|[Session](results-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
