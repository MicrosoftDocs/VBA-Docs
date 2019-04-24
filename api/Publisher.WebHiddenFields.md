---
title: WebHiddenFields object (Publisher)
keywords: vbapb10.chm4063231
f1_keywords:
- vbapb10.chm4063231
ms.prod: publisher
api_name:
- Publisher.WebHiddenFields
ms.assetid: 8ced4021-fa99-39dd-e880-b9793426871f
ms.date: 06/08/2017
localization_priority: Normal
---


# WebHiddenFields object (Publisher)

Represents hidden Web fields that allow a webpage to pass non-visible data to the web server when a webpage is submitted. The  **WebHiddenFields** object enables control of all the hidden fields attached to a Submit command button.
 


## Example

Use the  **HiddenFields** property to access hidden Web fields. This example adds a new hidden Web field to a new Submit command button.
 

 

```vb
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .HiddenFields.Add Name:="User", Value:="PowerUser" 
 End With 
 End With 
End Sub
```


## Methods



|Name|
|:-----|
|[Add](Publisher.WebHiddenFields.Add.md)|
|[Delete](Publisher.WebHiddenFields.Delete.md)|
|[Item](Publisher.WebHiddenFields.Item.md)|
|[Name](Publisher.WebHiddenFields.Name.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.WebHiddenFields.Application.md)|
|[Count](Publisher.WebHiddenFields.Count.md)|
|[Parent](Publisher.WebHiddenFields.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]