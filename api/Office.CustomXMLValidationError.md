---
title: CustomXMLValidationError object (Office)
keywords: vbaof11.chm307000
f1_keywords:
- vbaof11.chm307000
ms.prod: office
api_name:
- Office.CustomXMLValidationError
ms.assetid: 7f7ced9a-0878-9287-fe66-a7f0ffdc45b6
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLValidationError object (Office)

Represents a single validation error in a **CustomXMLValidationErrors** collection.


## Remarks

Validation errors can either be triggered when validating an operation against the schema, such as when adding a node, or triggered by the user when some condition is not met. For example, if a start date is later than an end date. 


## Example

The following example adds a custom part and then adds a child node to that part. Any errors that occur are added to the **CustomXMLValidationErrors** collection and then displayed in the Debug window.


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


## See also

- [CustomXMLValidationError object members](overview/library-reference/customxmlvalidationerror-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)
