---
title: Document.BuildingBlockInsert event (Word)
keywords: vbawd10.chm4001016
f1_keywords:
- vbawd10.chm4001016
ms.prod: word
api_name:
- Word.Document.BuildingBlockInsert
ms.assetid: 6c4b1f1f-da22-63b9-a3d9-21d7bedb4b5b
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.BuildingBlockInsert event (Word)

Occurs when you insert a building block into a document. .


## Syntax

_expression_.**BuildingBlockInsert'(**_Range_**, **_Name_**, **_Category_**, **_Type_**, **_Template_**)

 _expression_ An expression that returns a [Document](./Word.Document.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|Specifies the position where the building block is inserted.|
| _Name_|Required| **String**|Specifies the name of the building block.|
| _Category_|Required| **String**|Specifies the building block category.|
| _Type_|Required| **String**|Specifies the type of building block.|
| _Template_|Required| **String**|Specifies the name of the template that contains the building block.|

## Remarks

For information about using events with a  **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]