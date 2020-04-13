---
title: Application.DefaultSaveFormat property (Word)
keywords: vbawd10.chm158335040
f1_keywords:
- vbawd10.chm158335040
ms.prod: word
api_name:
- Word.Application.DefaultSaveFormat
ms.assetid: e15d8cc9-f6da-ccb0-784f-02fe9dc7ee6a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DefaultSaveFormat property (Word)

Returns or sets the default format that will appear in the **Save as type** box in the **Save As** dialog box. Read/write **String**.


## Syntax

_expression_. `DefaultSaveFormat`

 _expression_ An expression that represents a '[Application](Word.Application.md)' object.


## Remarks

The string used with this property is the file converter class name. The class names for internal Word formats are listed in the following table.



|**Word format**|**File converter class name**|
|:-----|:-----|
|Word Document|""|
|Document Template|"Dot"|
|Text Only|"Text"|
|Text Only with Line Breaks|"CRText"|
|MS-DOS Text|"8Text"|
|MS-DOS Text with Line Breaks|"8CRText"|
|Rich Text Format|"Rtf"|
|Unicode Text|"Unicode"|

Use the **[ClassName](Word.FileConverter.ClassName.md)** property of the **[FileConverter](Word.FileConverter.md)** object to determine the class name of an external file converter.


## Example

This example sets the Word document format as the default save format.


```vb
Application.DefaultSaveFormat = ""
```

This example returns the current setting that Word uses for saving a document.




```vb
Msgbox Application.DefaultSaveFormat
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]