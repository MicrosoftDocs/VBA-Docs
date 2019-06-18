---
title: WebHiddenFields.Name method (Publisher)
keywords: vbapb10.chm3997703
f1_keywords:
- vbapb10.chm3997703
ms.prod: publisher
api_name:
- Publisher.WebHiddenFields.Name
ms.assetid: 9dade2c9-6f6b-8686-90fa-a41c8bb6dfa2
ms.date: 06/18/2019
localization_priority: Normal
---


# WebHiddenFields.Name method (Publisher)

Returns a **String** that represents the name of a hidden web field for a web command button.


## Syntax

_expression_.**Name** (_Index_)

_expression_ A variable that represents a **[WebHiddenFields](Publisher.WebHiddenFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Long**|The index number of the hidden field.|

## Return value

String


## Example

This example creates a web command button with a hidden field, and then displays the field's name.

```vb
Sub GetHiddenWebFieldName() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, _ 
 Left:=100, Top:=100, Width:=100, _ 
 Height:=36).WebCommandButton.HiddenFields 
 .Add Name:="User", Value:="Power" 
 MsgBox "The name of the first hidden field is " & .Name(1) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]