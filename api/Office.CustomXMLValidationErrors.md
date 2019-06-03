---
title: CustomXMLValidationErrors object (Office)
keywords: vbaof11.chm308000
f1_keywords:
- vbaof11.chm308000
ms.prod: office
api_name:
- Office.CustomXMLValidationErrors
ms.assetid: 17c7b3dc-f4ba-b247-498d-48be197bbc91
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLValidationErrors object (Office)

Represents a collection of **[CustomXMLValidationError](Office.CustomXMLValidationError.md)** objects.


## Example

The following example adds a custom part, and then adds a child node to that part. Any errors that occur are added to the **CustomXMLValidationErrors** collection and then displayed in the Debug window.


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
      DeBug.Print("Error name: " & ValError.Name & " Error description: " & ValError.Text)  
   Next 
End If 
 
Exit Sub 
 
validation_error: 
   CustomXMLValidationErrors.Add(ValError.Name, ValError.Text)) 
Resume 

```


## See also

- [CustomXMLValidationErrors object members](overview/library-reference/customxmlvalidationerrors-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]