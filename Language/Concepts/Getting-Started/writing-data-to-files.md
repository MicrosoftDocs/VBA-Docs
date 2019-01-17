---
title: Writing data to files (VBA)
keywords: vbcn6.chm1010980
f1_keywords:
- vbcn6.chm1010980
ms.prod: office
ms.assetid: 019b8569-4ecf-e0bb-ce62-b2e79b2cf6dd
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing data to files

When working with large amounts of data, it is often convenient to write data to or read data from a file. The **[Open](../../reference/user-interface-help/open-statement.md)** statement lets you create and access files directly. **Open** provides three types of file access:

- Sequential access (**Input**, **Output**, and **Append** modes) is used for writing text files, such as error logs and reports.
    
- Random access (**Random** mode) is used to read and write data to a file without closing it. Random access files keep data in records, which makes it easy to locate information quickly.
    
- Binary access (**Binary** mode) is used to read or write to any byte position in a file, such as storing or displaying a bitmap image.
    
> [!NOTE] 
> The **Open** statement should not be used to open an application's own file types. For example, don't use **Open** to open a Word document, a Microsoft Excel spreadsheet, or a Microsoft Access database. Doing so will cause loss of file integrity and file corruption.

The following table shows the statements typically used when writing data to and reading data from files.

|Access type|Writing data|Reading data|
|:-----|:-----|:-----|
|Sequential|**Print #**, **Write #**|**Input #**|
|Random|**Put**|**Get**|
|Binary|**Put**|**Get**|

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]