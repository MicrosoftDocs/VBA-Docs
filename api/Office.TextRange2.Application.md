---
title: TextRange2.Application property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Application
ms.assetid: 3883561f-229b-92f9-eaea-83f00ac33f06
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.Application property (Office)

When used without an object qualifier, this property returns an **Application** object that represents the current instance of the Microsoft Office application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the **TextRange2** object. When used with an OLE **Automation** object, it returns the object's application. Read-only.


## Syntax

_expression_.**Application**

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


## Return value

Object


## Example

This example displays the name of the application that created each linked OLE object on page one of the active Publisher publication.


```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]