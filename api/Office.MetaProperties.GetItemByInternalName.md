---
title: MetaProperties.GetItemByInternalName method (Office)
keywords: vbaof11.chm274002
f1_keywords:
- vbaof11.chm274002
ms.prod: office
api_name:
- Office.MetaProperties.GetItemByInternalName
ms.assetid: 27c6bcd8-8631-1dbe-5df1-67c33b757c03
ms.date: 01/18/2019
localization_priority: Normal
---


# MetaProperties.GetItemByInternalName method (Office)

Gets a property's value specifying its name as opposed to its index value.


## Syntax

_expression_.**GetItemByInternalName**(_InternalName_)

_expression_ An expression that returns a **[MetaProperties](Office.MetaProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _InternalName_|Required|**String**|Contains the name of the property.|

## Return value

MetaProperty


## Remarks

Metadata is information about a document that can be used to identify particular documents, search document content, build rich content dynamically, and other similar operations without the need to open the document. Metadata can be stored in a document and as properties on a Microsoft SharePoint Foundation server.


## Example

In the following example, a **MetaProperties** object is passed to a validation function. The function then retrieves the value of one of its properties and assigns it to a **MetaProperty** object. Finally, the property is validated and the results are returned.


```vb
Function ValidateMetaProperty(ByVal objMetaProperty As MetaProperty) As String 
Dim objMetaProperty As MetaProperty 
Dim result As String 
 
objMetaProperty = objMetaProperty.GetItemByInternalName("type") 
result = objMetaProperty.Validate 
 
ValidateMetaProperty = result 
End Function
```


## See also

- [MetaProperties object members](overview/Library-Reference/metaproperties-members-office.md)






[!include[Support and feedback](~/includes/feedback-boilerplate.md)]