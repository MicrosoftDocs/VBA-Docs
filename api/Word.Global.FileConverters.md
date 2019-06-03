---
title: Global.FileConverters property (Word)
keywords: vbawd10.chm163119121
f1_keywords:
- vbawd10.chm163119121
ms.prod: word
api_name:
- Word.Global.FileConverters
ms.assetid: 1112fe8e-5a32-bbef-d02f-f17f0603f539
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.FileConverters property (Word)

Returns a  **[FileConverters](Word.fileconverters.md)** collection that represents all the file converters available to Microsoft Word. Read-only.


## Syntax

_expression_. `FileConverters`

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]