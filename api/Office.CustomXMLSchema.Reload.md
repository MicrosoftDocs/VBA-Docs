---
title: CustomXMLSchema.Reload method (Office)
keywords: vbaof11.chm291005
f1_keywords:
- vbaof11.chm291005
ms.prod: office
api_name:
- Office.CustomXMLSchema.Reload
ms.assetid: 963b941a-0b93-fc02-c150-747975005561
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLSchema.Reload method (Office)

Reloads a schema from a file.


## Syntax

_expression_.**Reload**

_expression_ An expression that returns a **[CustomXMLSchema](Office.CustomXMLSchema.md)** object.


## Remarks

Typically, this method is used to update the location of the schema or to determine if the schema is still valid. It is also useful for reloading a schema that frequently changes. If this action is attempted on a schema in a collection that is already validated or tied to a data stream, the operation is not performed and an error message is displayed.


## Example

The following example specifies the location of the schema and then reloads it.


```vb
Dim objCustomXMLSchema As  CustomXMLSchema 
Dim strSchemaLocation As String 
' Set the location of the schema.. 
objCustomXMLSchema.Location = "c:\mySchema.xsd" 
 
' Reload the schema. 
objCustomXMLSchema.Reload 

```


## See also

- [CustomXMLSchema object members](overview/library-reference/customxmlschema-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]