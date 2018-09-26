---
title: TextRange2.Sentences Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Sentences
ms.assetid: 236196a7-97b3-f3d5-b483-c42bc60bd9ed
ms.date: 06/08/2017
---


# TextRange2.Sentences Property (Office)

Returns a  **TextRange2** object that represents the specified subset of text sentences. Read-only.


## Syntax

 _expression_. `Sentences`( `_Start_`, `_Length_` )

 _expression_ An expression that returns a [TextRange2](./Office.TextRange2.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first sentence in the returned range.|
| _Length_|Optional|**Long**|The number of sentences to be returned.|

### Return value

TextRange2


## Remarks

If both Start and Length are omitted, the returned range starts with the first sentence and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one sentence.

If Length is specified but Start is omitted, the returned range starts with the first sentence in the specified range.

If Start is greater than the number of sentences in the specified text, the returned range starts with the last sentence in the specified range.

If Length is greater than the number of sentences from the specified starting sentence to the end of the text, the returned range contains all those sentences.


## Example

This example formats as bold the second sentence in the second paragraph in shape two on slide one in the active PowerPoint presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _ 
 .TextFrame.TextRange2.Paragraphs(2).Sentences(2).Font _ 
 .Bold = True 
 
```


## See also


[TextRange2 Object](Office.TextRange2.md)



[TextRange2 Object Members](./overview/Library-Reference/textrange2-members-office.md)

