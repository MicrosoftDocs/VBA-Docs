---
title: PageBackground object (Publisher)
keywords: vbapb10.chm8191999
f1_keywords:
- vbapb10.chm8191999
ms.prod: publisher
api_name:
- Publisher.PageBackground
ms.assetid: 647f5a84-0971-2f69-d281-c9ab402968a4
ms.date: 06/01/2019
localization_priority: Normal
---


# PageBackground object (Publisher)

Represents the background of a page.
 
## Remarks

Use the **[Background](publisher.page.background.md)** property of a **Page** object to return a **PageBackground** object. 

Use the **Exists** property to determine if a background already exists for the specified **Page** object. 

Use the **Fill** property to return a **FillFormat** object. 

Use the **Delete** method to delete a background for the specified page. 

## Example

The following example creates a **PageBackground** object and sets it to the background of the first page of the active document.

```vb
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
 
```

<br/>

The following example builds upon the previous example. First a **PageBackground** object is created and set to the background of the first page of the active document. Next, a test is made to check if a background exists for the page already. If not, one is created by calling the **Create** method of the **PageBackground** object.

```vb
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
If objPageBackground.Exists = False Then 
 objPageBackground.Create 
End If 
 
```

<br/>

The following example builds upon the previous example. First a **PageBackground** object is created and set to the background of the first page of the active document. Next, a test is made to check if a background exists for the page already. If not, one is created by calling the **Create** method of the **PageBackground** object. A **FillFormat** object is returned by using the **Fill** property of the **PageBackground** object. A few of the available properties of the **FillFormat** object are then set.

```vb
Dim objPageBackground As PageBackground 
Dim objFillFormat As FillFormat 
 
Set objPageBackground = ActiveDocument.Pages(1).Background 
If objPageBackground.Exists = False Then 
 objPageBackground.Create 
End If 
 
Set objFillFormat = objPageBackground.Fill 
With objFillFormat 
 .BackColor.RGB = RGB(Red:=0, GReen:=155, Blue:=99) 
 .ForeColor.RGB = RGB(Red:=155, GReen:=234, Blue:=0) 
 .TwoColorGradient msoGradientDiagonalDown, 4 
End With 
 
```

<br/>

The following example deletes the background of the first page in the active document. This example assumes that the specified page has an existing background. A run-time error occurs if the page does not contain a background.

```vb
ActiveDocument.Pages(1).Background.Delete
```


## Methods

- [Create](Publisher.PageBackground.Create.md)
- [Delete](Publisher.PageBackground.Delete.md)

## Properties

- [Application](Publisher.PageBackground.Application.md)
- [Exists](Publisher.PageBackground.Exists.md)
- [Fill](Publisher.PageBackground.Fill.md)
- [Parent](Publisher.PageBackground.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]