---
title: CustomXMLSchemaCollection.AddCollection method (Office)
keywords: vbaof11.chm292006
f1_keywords:
- vbaof11.chm292006
ms.prod: office
api_name:
- Office.CustomXMLSchemaCollection.AddCollection
ms.assetid: d3b49c57-9a5b-9b5b-0003-d09240d227c1
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLSchemaCollection.AddCollection method (Office)

Adds an existing schema collection to the current schema collection. 


## Syntax

_expression_.**AddCollection**(_SchemaCollection_)

_expression_ An expression that returns a **[CustomXMLSchemaCollection](Office.CustomXMLSchemaCollection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SchemaCollection_|Required|**CustomXMLSchemaCollection**|Represents a collection of schemas to be imported into the current schema collection.|

## Remarks

If there is a conflict between namespaces while importing the collection, for example if x.xsd is already linked to "urn:invoice:namespace" but the incoming collection has z.xsd for the same namespace, the incoming collection wins.


## Example

The following example receives the target schema collection and incoming schema collection arguments and then adds the one collection to the other.


```vb
Sub AddSchema(objTargetCustomXMLSchemaCollection As CustomXMLSchemaCollection, _ 
  objTargetCustomXMLSchemaCollection As CustomXMLSchemaCollection) 
 
    ' Adds a schema collection to another schema the collection. 
    objTargetCustomXMLSchemaCollection.AddCollection(objIncomingCustomXMLSchemaCollection) 
                
End Sub
```


## See also

- [CustomXMLSchemaCollection object members](overview/library-reference/customxmlschemacollection-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]