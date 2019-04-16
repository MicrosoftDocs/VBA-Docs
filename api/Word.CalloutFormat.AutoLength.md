---
title: CalloutFormat.AutoLength property (Word)
keywords: vbawd10.chm163905639
f1_keywords:
- vbawd10.chm163905639
ms.prod: word
api_name:
- Word.CalloutFormat.AutoLength
ms.assetid: 345f77e7-0043-9c4f-e981-18f370314db1
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.AutoLength property (Word)

 **MsoTrue** to automatically sets the length of the callout line. Read-only **MsoTriState**.


## Syntax

_expression_.**AutoLength**

_expression_ Required. A variable that represents a '[CalloutFormat](Word.CalloutFormat.md)' object.


## Remarks

Use the  **AutomaticLength** method to set this property to **msoTrue**, and use the **CustomLength** method to set this property to **msoFalse**.


## Example

This example creates a new document and adds a callout to the new document, and then sets the length of the callout manually.


```vb
Sub AutoCalloutLength() 
 Dim docNew As Document 
 Dim shpCallout As Shape 
 Set docNew = Documents.Add 
 Set shpCallout = docNew.Shapes.AddCallout(Type:=msoCalloutFour, _ 
 Left:=15, Top:=15, Width:=150, Height:=200) 
 With shpCallout.Callout 
 If .AutoLength = msoTrue then 
 .CustomLength 50 
 End If 
 End With 
End Sub
```


## See also


[CalloutFormat Object](Word.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]