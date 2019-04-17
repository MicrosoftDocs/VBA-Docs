---
title: FileSystemObject object
keywords: vblr6.chm2181927
f1_keywords:
- vblr6.chm2181927
ms.prod: office
api_name:
- Office.FileSystemObject
ms.assetid: 7ad2dad3-c6d8-90a6-77a5-c712da8316f3
ms.date: 11/12/2018
localization_priority: Priority
---

# FileSystemObject object

Provides access to a computer's file system.

## Syntax

**Scripting.FileSystemObject**

## Remarks

The following code illustrates how the **FileSystemObject** object is used to return a **[TextStream](textstream-object.md)** object that can be read from or written to:

```vb
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\testfile.txt", True)
a.WriteLine("This is a test.")
a.Close
```

In the example code:

- The **[CreateObject](createobject-function.md)** function returns the **FileSystemObject** (`fs`). 
- The **[CreateTextFile](createtextfile-method.md)** method creates the file as a **TextStream** object (`a`).
- The **[WriteLine](writeline-method.md)** method writes a line of text to the created text file. 
- The **[Close](close-method-textstream-object.md)** method flushes the buffer and closes the file.

## Methods

|Method|Description|
|:-----|:----------|
|[BuildPath](buildpath-method.md)|Appends a name to an existing path. |
|[CopyFile](copyfile-method.md)|Copies one or more files from one location to another. |
|[CopyFolder](copyfolder-method.md)|Copies one or more folders from one location to another. |
|[CreateFolder](createfolder-method.md)|Creates a new folder. |
|[CreateTextFile](createtextfile-method.md)|Creates a text file and returns a TextStream object that can be used to read from, or write to the file. |
|[DeleteFile](deletefile-method.md)|Deletes one or more specified files. |
|[DeleteFolder](deletefolder-method.md)|Deletes one or more specified folders. |
|[DriveExists](driveexists-method.md)|Checks if a specified drive exists. |
|[FileExists](fileexists-method.md)|Checks if a specified file exists. |
|[FolderExists](folderexists-method.md)|Checks if a specified folder exists. |
|[GetAbsolutePathName](getabsolutepathname-method.md)|Returns the complete path from the root of the drive for the specified path. |
|[GetBaseName](getbasename-method.md)|Returns the base name of a specified file or folder. |
|[GetDrive](getdrive-method.md)|Returns a Drive object corresponding to the drive in a specified path. |
|[GetDriveName](getdrivename-method.md)|Returns the drive name of a specified path. |
|[GetExtensionName](getextensionname-method.md)|Returns the file extension name for the last component in a specified path. |
|[GetFile](getfile-method.md)|Returns a File object for a specified path. |
|[GetFileName](getfilename-method-visual-basic-for-applications.md)|Returns the file name or folder name for the last component in a specified path. |
|[GetFolder](getfolder-method.md)|Returns a Folder object for a specified path. |
|[GetParentFolderName](getparentfoldername-method.md)|Returns the name of the parent folder of the last component in a specified path. |
|[GetSpecialFolder](getspecialfolder-method.md)|Returns the path to some of Windows' special folders. |
|[GetTempName](gettempname-method.md)|Returns a randomly generated temporary file or folder. |
|[Move](move-method-filesystemobject-object.md)|Moves a specified file or folder from one location to another. |
|[MoveFile](movefile-method.md)|Moves one or more files from one location to another. |
|[MoveFolder](movefolder-method.md)|Moves one or more folders from one location to another. |
|[OpenAsTextStream](openastextstream-method.md)|Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file. |
|[OpenTextFile](opentextfile-method.md)|Opens a file and returns a TextStream object that can be used to access the file. |
|[WriteLine](writeline-method.md)|Writes a specified string and new-line character to a TextStream file. |

## Properties

|Property|Description|
|:-------|:----------|
|[Drives](drives-property.md)|Returns a collection of all Drive objects on the computer. |
|[Name](name-property-filesystemobject-object.md)|Sets or returns the name of a specified file or folder. |
|[Path](path-property-filesystemobject-object.md)|Returns the path for a specified file, folder, or drive. |
|[Size](size-property-filesystemobject-object.md)|For files, returns the size, in bytes, of the specified file; for folders, returns the size, in bytes, of all files and subfolders contained in the folder. |
|[Type](type-property-filesystemobject-object.md)|Returns information about the type of a file or folder (for example, for files ending in .TXT, "Text Document" is returned). |

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Properties (Visual Basic for Applications)](../properties-visual-basic-for-applications.md)
- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
