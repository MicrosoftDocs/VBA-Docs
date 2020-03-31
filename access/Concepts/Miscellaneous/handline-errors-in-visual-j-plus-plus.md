---
title: Handline errors in Visual J++
ms.prod: access
ms.assetid: 100fca9d-38e3-e31c-71ce-29c928fbef88
ms.date: 06/08/2019
localization_priority: Normal
---


# Handline errors in Visual J++

**Applies to:** Access 2013 | Access 2016

Handle ADO errors in your Microsoft Visual J++ applications using a **try catch** block. Once an error has been thrown, you can iterate through the collection, successively handling each error. The following Visual J++ example shows a console application that deliberately causes an error.

When the **catch** block is activated, it calls the PrintProviderError function to display the errors. The PrintProviderError function iterates through the **Errors** collection and sends a line to the standard output device that describes each error in the collection.

```cs
 
// BeginErrorExampleVJ 
/** 
 * This class can take a variable number of parameters on the command 
 * line. Program execution begins with the main() method. The class 
 * constructor is not invoked unless an object of type 'Class1' 
 * created in the main() method. 
 */ 
 
import com.ms.wfc.data.*; 
import java.io.* ; 
 
public class ErrorExample 
{ 
 /** 
 * The main entry point for the application. 
 * 
 * @param args Array of parameters passed to the application 
 * via the command line. 
 */ 
 public static void main (String[] args) 
 { 
 DescriptionX(); 
 System.exit(0); 
 } 
 
 static void DescriptionX() 
 { 
 BufferedReader in = new 
 BufferedReader(new InputStreamReader(System.in)); 
 
 // Define ADO Objects. 
 Connection cnConn1 = null; 
 
 try 
 { 
 // Create an error by trying to 
 // Open a database that doesn't exist. 
 cnConn1 = new Connection(); 
 cnConn1.open("nothing"); 
 } 
 catch( AdoException ae ) 
 { 
 // Notify user of any errors that result from ADO. 
 PrintProviderError(cnConn1); 
 } 
 
 try 
 { 
 System.out.println("\nPress <Enter> key to continue."); 
 in.readLine(); 
 } 
 // System read requires this catch. 
 catch( java.io.IOException je) 
 { 
 PrintIOError(je); 
 } 
 } 
 
 // PrintProviderError Function 
 static void PrintProviderError( Connection Cnn1 ) 
 { 
 // Print Provider errors from Connection object. 
 // ErrItem is an item object in the Connections Errors collection. 
 com.ms.wfc.data.Error ErrItem = null; 
 long nCount = 0; 
 int i = 0; 
 
 nCount = Cnn1.getErrors().getCount(); 
 
 // If there are any errors in the collection, print them. 
 if( nCount > 0); 
 { 
 // Collection ranges from 0 to nCount - 1 
 for (i = 0; i< nCount; i++) 
 { 
 ErrItem = Cnn1.getErrors().getItem(i); 
 System.out.println("\t Error number: " + ErrItem.getNumber() 
 + "\t" + ErrItem.getDescription() ); 
 } 
 } 
 } 
 
 // PrintIOError Function 
 static void PrintIOError( java.io.IOException je) 
 { 
 System.out.println("Error \n"); 
 System.out.println("\tSource = " + je.getClass() + "\n"); 
 System.out.println("\tDescription = " + je.getMessage() + "\n"); 
 } 
} 
// EndErrorExampleVJ 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]