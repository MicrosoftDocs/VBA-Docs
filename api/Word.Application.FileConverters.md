---
title: Application.FileConverters property (Word)
keywords: vbawd10.chm158334993
f1_keywords:
- vbawd10.chm158334993
ms.prod: word
api_name:
- Word.Application.FileConverters
ms.assetid: 90f58ceb-6fb3-ee48-9637-effe39163888
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FileConverters property (Word)

Returns a  **[FileConverters](Word.fileconverters.md)** collection that represents all the file converters available to Microsoft Word. Read-only.


## Syntax

_expression_. `FileConverters`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the path of the WordPerfect 5.0 file converter.


```vb
MsgBox FileConverters("WrdPrfctDOS50").Path
```

This example displays a message that indicates whether the third converter in the FileConverters collection can save files.




```vb
If FileConverters(3).CanSave = True Then 
 MsgBox FileConverters(3).FormatName & " can save files" 
Else 
 MsgBox FileConverters(3).FormatName & " cannot save files" 
End If
```

This example displays the name of the last file converter.




```vb
Dim fcTemp As FileConverter 
 
Set fcTemp = FileConverters(FileConverters.Count) 
MsgBox "The file name extensions for " & fcTemp.FormatName & _ 
 " files are: " & fcTemp.Extensions
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]