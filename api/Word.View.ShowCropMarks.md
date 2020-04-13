---
title: View.ShowCropMarks property (Word)
keywords: vbawd10.chm161808440
f1_keywords:
- vbawd10.chm161808440
ms.prod: word
api_name:
- Word.View.ShowCropMarks
ms.assetid: bc6db5f2-a9e4-5c0a-7e1a-43a93620f12b
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowCropMarks property (Word)

Returns or sets a  **Boolean** that represents whether to show crop marks in the corners of pages to indicate where margins are located. Read/write.


## Syntax

_expression_. `ShowCropMarks`

 _expression_ An expression that returns a [View](./Word.View.md) object.


## Remarks

Displaying crop marks does not allow a user to change the margins by dragging the crop marks. Crop marks are only displayed to indicate where margins are located in the page. This property corresponds to the **Crop marks** check box in the **Advanced** tab of the **Word Options** dialog box.


> [!NOTE] 
> Crop marks are shown by default in East Asian languages and are off by default in all other languages.


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]