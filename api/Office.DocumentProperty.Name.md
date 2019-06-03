---
title: DocumentProperty.Name property (Office)
keywords: vbaof11.chm250005
f1_keywords:
- vbaof11.chm250005
ms.prod: office
api_name:
- Office.DocumentProperty.Name
ms.assetid: b609c38e-71ca-e019-9852-fc7811dc798f
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty.Name property (Office)

Gets or sets the name of a document property. Read/write.


## Syntax

_expression_.**Name**(_lcid_, _pbstrRetVal_)

_expression_ A variable that represents a **[DocumentProperty](Office.DocumentProperty.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lcid_|Required|**Long**|Represents the language identifier.|
| _pbstrRetVal_|Required|**String**|Represents the return value for the property.|

## Return value

String


## Remarks

A **DocumentProperty** object represents a custom or built-in document property of a container document.


## Example

This example displays the name, type, and value of a document property. You must pass a valid **DocumentProperty** object to the procedure.

```vb
Sub DisplayPropertyInfo(dp As DocumentProperty) 
 MsgBox "value = " & dp.Value & Chr(13) & _ 
 "type = " & dp.Type & Chr(13) & _ 
 "name = " & dp.Name 
End Sub
```


## See also

- [DocumentProperty object members](overview/library-reference/documentproperty-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]