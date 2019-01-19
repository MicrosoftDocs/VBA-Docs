---
title: MetaProperty object (Office)
keywords: vbaof11.chm275000
f1_keywords:
- vbaof11.chm275000
ms.prod: office
api_name:
- Office.MetaProperty
ms.assetid: 4379d183-9b80-92d8-1dd0-ac9be400e366
ms.date: 01/18/2019
localization_priority: Normal
---


# MetaProperty object (Office)

Represents a single property in a collection of properties describing the metadata stored in a document.


## Example

In the following example, a **[MetaProperties](office.metaproperties.md)** object is passed to a validation function. The function then validates the value of a single property represented by its index and returns the result.


```vb
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps(1).Validate 
ValidateMetaProperty = result 
End Function 

```


## See also

- [MetaProperty object members](overview/Library-Reference/metaproperty-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]