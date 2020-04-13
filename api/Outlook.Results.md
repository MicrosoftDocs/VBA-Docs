---
title: Results object (Outlook)
keywords: vbaol11.chm3012
f1_keywords:
- vbaol11.chm3012
ms.prod: outlook
api_name:
- Outlook.Results
ms.assetid: 59057f6f-8f6d-eed0-c945-240b9593b7ea
ms.date: 06/08/2017
localization_priority: Normal
---


# Results object (Outlook)

Contains data and results returned by the  **[Search](Outlook.Search.md)** object and the **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method.


## Remarks

The **Results** object contains properties and methods that allow you to view and manipulate data. For example the **[GetNext](Outlook.Results.GetNext.md)**, **[GetPrevious](Outlook.Results.GetPrevious.md)**, **[GetFirst](Outlook.Results.GetFirst.md)**, and **[GetLast](Outlook.Results.GetLast.md)** methods allow you to search through the results and view the data by field. The **[Sort](Outlook.Results.Sort.md)** method allows you to sort the data.

Use the  **SearchObject.Results** property to return a **Results** object.


## Example

The following event procedure stores the results of a search in a variable named objRsts and displays the results of the search in the Immediate window.


```vb
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 
 
 Dim objRsts As Outlook.Results 
 
 MsgBox "The search " & SearchObject.Tag & _ 
 
 "has completed. The scope of the search was " & _ 
 
 SearchObject.Scope & "." 
 
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



|Name|
|:-----|
|[ItemAdd](Outlook.Results.ItemAdd.md)|
|[ItemChange](Outlook.Results.ItemChange.md)|
|[ItemRemove](Outlook.Results.ItemRemove.md)|

## Methods



|Name|
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



|Name|
|:-----|
|[Application](Outlook.Results.Application.md)|
|[Class](Outlook.Results.Class.md)|
|[Count](Outlook.Results.Count.md)|
|[DefaultItemType](Outlook.Results.DefaultItemType.md)|
|[Parent](Outlook.Results.Parent.md)|
|[Session](Outlook.Results.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]