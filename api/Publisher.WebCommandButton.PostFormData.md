---
title: WebCommandButton.PostFormData property (Publisher)
keywords: vbapb10.chm3932176
f1_keywords:
- vbapb10.chm3932176
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.PostFormData
ms.assetid: d04e3172-0d96-856f-af63-341031d92291
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.PostFormData property (Publisher)

Returns or sets an **[MsoTriState](office.msotristate.md)** constant indicating whether the specified web command button control uses the Microsoft Visual Basic **Get** or **Post** method when submitting form data to a web server. Read/write.


## Syntax

_expression_.**PostFormData**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Return value

MsoTriState


## Remarks

The property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The control uses the Visual Basic **Get** method to submit form data.|
| **msoTrue**|The control uses the Visual Basic **Post** method to submit form data. The default value.|

This property is ignored for **Reset** command buttons.


## Example

This example creates a web form **Submit** command button and sets the script path and file name to run when a user chooses the button. The example also specifies that the web form should use the Visual Basic **Get** method to submit form data.

```vb
Dim shpNew As Shape 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36) 
 
With shpNew.WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "https://www.tailspintoys.com/" _ 
 & "scripts/ispscript.cgi" 
 .PostFormData = msoFalse 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]