---
title: CustomXMLSchema.Delete method (Office)
keywords: vbaof11.chm291004
f1_keywords:
- vbaof11.chm291004
ms.prod: office
api_name:
- Office.CustomXMLSchema.Delete
ms.assetid: bdd79a25-7f2f-c810-13b0-9d7dc34e9a3d
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLSchema.Delete method (Office)

Deletes the specified schema from the **CustomXMLSchema** collection.


## Syntax

_expression_.**Delete**

_expression_ An expression that returns a **[CustomXMLSchema](Office.CustomXMLSchema.md)** object.


## Remarks

If this operation is attempted on a schema in a collection that is already validated or attached to a data stream, the operation is not performed and an error message is displayed.


## Example

The following example adds a schema to the collection and then deletes the schema.


```vb
Sub DeleteSchema() 
    On Error GoTo Err 
 
    Dim objCustomXMLSchemaCollection As CustomXMLSchemaCollection 
    Dim objCustomXMLSchema As  CustomXMLSchema 
 
    ' Adds a schema to the collection. 
    objCustomXMLSchema.Add("urn:invoice:namespace")  
 
    ... 
 
    ' Deletes the schema. 
    objCustomXMLSchema.Delete 
      
    Exit Sub 
                 
' Exception handling. Show the message and resume. 
Err: 
        MsgBox (Err.Description) 
        Resume Next 
End Sub
```


## See also

- [CustomXMLSchema object members](overview/library-reference/customxmlschema-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]