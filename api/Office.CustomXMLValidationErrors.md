---
title: CustomXMLValidationErrors Object (Office)
keywords: vbaof11.chm308000
f1_keywords:
- vbaof11.chm308000
ms.prod: office
api_name:
- Office.CustomXMLValidationErrors
ms.assetid: 17c7b3dc-f4ba-b247-498d-48be197bbc91
ms.date: 06/08/2017
---


# CustomXMLValidationErrors Object (Office)

Represents a collection of  **CustomXMLValidationError** objects.


## Example

The following example adds a custom part and then adds a child node to that part. Any errors that occur are added to the  **CustomXMLValidationErrors** collection and then displayed in the Debug window.


```vb
Dim ValErrors As CustomXMLValidationErrors 
Dim ValError As CustomXMLValidationError 
Dim cxp1 As CustomXMLPart 
Dim intError As Integer 
 
On Error Go To validation_error 
 
 With ActiveDocument 
 
    ' Add and populate a custom xml part 
    set cxp1 = .CustomXMLParts.Add "<invoice>" 
 
    ' Add a node 
    cxp1.AddNode "<quantity>", "supplier", "urn:invoice:namespace" 
 
 End With 
 
If ValErrors.Count > 0 then 
   For Each ValError In ValErrors 
      DeBug.Print("Error name: " &amp; ValError.Name &amp; " Error description: " &amp; ValError.Text)  
   Next 
End If 
 
Exit Sub 
 
validation_error: 
   CustomXMLValidationErrors.Add(ValError.Name, ValError.Text)) 
Resume 

```


## Methods



|**Name**|
|:-----|
|[Add](Office.CustomXMLValidationErrors.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.CustomXMLValidationErrors.Application.md)|
|[Count](Office.CustomXMLValidationErrors.Count.md)|
|[Creator](Office.CustomXMLValidationErrors.Creator.md)|
|[Item](Office.CustomXMLValidationErrors.Item.md)|
|[Parent](Office.CustomXMLValidationErrors.Parent.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
