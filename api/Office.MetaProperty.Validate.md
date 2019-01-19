---
title: MetaProperty.Validate method (Office)
keywords: vbaof11.chm275007
f1_keywords:
- vbaof11.chm275007
ms.prod: office
api_name:
- Office.MetaProperty.Validate
ms.assetid: e8037c82-a9bd-936f-fbf1-03c35d83685b
ms.date: 01/18/2019
localization_priority: Normal
---


# MetaProperty.Validate method (Office)

Validates a **MetaProperty** object representing a single property value according to a schema.


## Syntax

_expression_.**Validate**

_expression_ An expression that returns a **[MetaProperty](Office.MetaProperty.md)** object.


## Return value

String


## Remarks

If the property is invalid, the test fails and an error message is returned. The schema used for validation is stored as part of the document's Microsoft SharePoint Foundation profile.


## Example

In the following example, a **[MetaProperties](Office.MetaProperties.md)** object is passed to a validation function. The function then validates the value of a single property represented by its index and returns the result.


```vb
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps(1).Validate 
ValidateMetaProperty = result 
End Function
```


## See also

- [MetaProperty object members](overview/Library-Reference/metaproperty-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]