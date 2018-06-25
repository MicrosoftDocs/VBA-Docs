---
title: CustomXMLSchemaCollection Object (Office)
keywords: vbaof11.chm306000
f1_keywords:
- vbaof11.chm306000
ms.prod: office
api_name:
- Office.CustomXMLSchemaCollection
ms.assetid: 0ce1fe79-4287-303a-4205-586d8e116731
ms.date: 06/08/2017
---


# CustomXMLSchemaCollection Object (Office)

Represents a collection of  **CustomXMLSchema** objects attached to a data stream.


## Example

The following example adds a  **CustomXMLSchema** object to a **CustomXMLSchemaCollection** object.


```vb
Dim SchemaCollection As CustomXMLSchemaCollection 
 
SchemaCollection.Add "http://tempuri.org/XMLSchema.xsd"
```


## Methods



|**Name**|
|:-----|
|[Add](Office.CustomXMLSchemaCollection.Add.md)|
|[AddCollection](Office.CustomXMLSchemaCollection.AddCollection.md)|
|[Validate](Office.CustomXMLSchemaCollection.Validate.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.CustomXMLSchemaCollection.Application.md)|
|[Count](Office.CustomXMLSchemaCollection.Count.md)|
|[Creator](Office.CustomXMLSchemaCollection.Creator.md)|
|[Item](Office.CustomXMLSchemaCollection.Item.md)|
|[NamespaceURI](Office.CustomXMLSchemaCollection.NamespaceURI.md)|
|[Parent](Office.CustomXMLSchemaCollection.Parent.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
