---
title: DocumentProperties.Add method (Office)
keywords: vbaof11.chm250014
f1_keywords:
- vbaof11.chm250014
ms.prod: office
api_name:
- Office.DocumentProperties.Add
ms.assetid: 80738562-8b0b-33f1-3dfa-0d66b1844ef7
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperties.Add method (Office)

Creates a new custom document property. You can add a new document property only to the custom **DocumentProperties** collection.


## Syntax

_expression_.**Add** (_Name_, _LinkToContent_, _Type_, _Value_, _LinkSource_)

_expression_ Required. A variable that represents a **[DocumentProperties](Office.DocumentProperties.md)** object. The custom **DocumentProperties** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The string of the **[Name](office.documentproperty.name.md)** of the property.|
| _LinkToContent_|Required|**Boolean**|Specifies whether the **[LinkToContent](office.documentproperty.linktocontent.md)** property is linked to the contents of the container document. If this argument is **True**, the _LinkSource_ argument is required; if it's **False**, the _Value_ argument is required.|
| _Type_|Optional|**Variant**|The data type of the **[Type](office.documentproperty.type.md)** property. Can be one of the following **[MsoDocProperties](office.msodocproperties.md)** constants: **msoPropertyTypeBoolean**, **msoPropertyTypeDate**, **msoPropertyTypeFloat**, **msoPropertyTypeNumber**, or **msoPropertyTypeString**.|
| _Value_|Optional|**Variant**|The data value of the **[Value](office.documentproperty.value.md)** property, if it's not linked to the contents of the container document. The value is converted to match the data type specified by the _Type_ argument, and if it can't be converted, an error occurs. If _LinkToContent_ is **True**, the argument is ignored, and the new document property is assigned a default value until the linked property values are updated by the container application (usually when the document is saved).|
| _LinkSource_|Optional|**Variant**|Ignored if _LinkToContent_ is **False**. The source of the **[LinkSource](office.documentproperty.linksource.md)** property. The container application determines what types of source linking you can use. For example, DDE links use the "Server\|Document!Item" syntax.|

<br/>

## Remarks

If you add a custom document property to the **DocumentProperties** collection that's linked to a given value in an Office document, you must save the document to see the change to the **DocumentProperty** object.

## Example

This example, which is designed to run in Word, adds three custom document properties to the **DocumentProperties** collection.

```vb
With ActiveDocument.CustomDocumentProperties 
    .Add Name:="LastModifiedBy", _ 
        LinkToContent:=True, _ 
        Type:=msoPropertyTypeString, _ 
        LinkSource:=Author
    .Add Name:="CustomNumber", _ 
        LinkToContent:=False, _ 
        Type:=msoPropertyTypeNumber, _ 
        Value:=1000 
    .Add Name:="CustomString", _ 
        LinkToContent:=False, _ 
        Type:=msoPropertyTypeString, _ 
        Value:="This is a custom property." 
    .Add Name:="CustomDate", _ 
        LinkToContent:=False, _ 
        Type:=msoPropertyTypeDate, _ 
        Value:=Date 
End With
```



## See also

- [DocumentProperties object members](overview/library-reference/documentproperties-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
