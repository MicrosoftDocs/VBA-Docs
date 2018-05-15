---
title: WebCommandButton Object (Publisher)
keywords: vbapb10.chm3997695
f1_keywords:
- vbapb10.chm3997695
ms.prod: publisher
api_name:
- Publisher.WebCommandButton
ms.assetid: 86605945-eca1-ab80-1a1a-f8a5977d9282
ms.date: 06/08/2017
---


# WebCommandButton Object (Publisher)

Represents a Web command button control. The  **WebCommandButton** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create new Web command button. Use the **[WebCommandButton](Publisher.Shape.WebCommandButton.md)** property to access a Web command button control shape. This example creates a Web form Submit command button and sets the script path and file name to run when a user clicks the button.
 

 

```
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" _ 
 &amp; "scripts/ispscript.cgi" 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[ActionURL](Publisher.WebCommandButton.ActionURL.md)|
|[Application](Publisher.WebCommandButton.Application.md)|
|[ButtonText](Publisher.WebCommandButton.ButtonText.md)|
|[ButtonType](Publisher.WebCommandButton.ButtonType.md)|
|[DataFileFormat](Publisher.WebCommandButton.DataFileFormat.md)|
|[DataFileName](Publisher.WebCommandButton.DataFileName.md)|
|[DataRetrievalMethod](Publisher.WebCommandButton.DataRetrievalMethod.md)|
|[EmailAddress](Publisher.WebCommandButton.EmailAddress.md)|
|[EmailSubject](Publisher.WebCommandButton.EmailSubject.md)|
|[HiddenFields](Publisher.WebCommandButton.HiddenFields.md)|
|[Parent](Publisher.WebCommandButton.Parent.md)|
|[PostFormData](Publisher.WebCommandButton.PostFormData.md)|

