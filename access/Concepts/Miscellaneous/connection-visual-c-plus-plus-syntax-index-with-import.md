---
title: Connection (Visual C++ syntax index with import)
ms.prod: access
ms.assetid: 3217a7d7-1c70-89f7-74a4-172371521358
ms.date: 10/12/2018
localization_priority: Normal
---


# Connection (Visual C++ syntax index with #import)

**Applies to:** Access 2013 | Access 2016

## Methods

```cs
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclose_HV10294173.xml( ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcnnexecute_HV10294345.xml( _bstr_t CommandText , VARIANT * 
 RecordsAffected , long Options ); 
 
long Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthbegintrans_HV10294108.xml( ); 
HRESULT CommitTrans( ); 
HRESULT RollbackTrans( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcnnopen_HV10294563.xml( _bstr_t ConnectionString , _bstr_t UserID , 
 _bstr_t Password , long Options ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthopenschema_HV10294568.xml( enum SchemaEnum Schema , const 
 _variant_t & Restrictions  = vtMissing, const _variant_t & 
 SchemaID  =vtMissing); 

```

## Properties

```cs
_bstr_t GetConnectionString( ); 
void PutConnectionString( _bstr_t pbstr ); 
__declspec(property(get=GetConnectionString,put=PutConnectionString)) 
 _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproconnectionstring_HV10294218.xml; 
 
long GetCommandTimeout( ); 
void PutCommandTimeout( long plTimeout ); 
__declspec(property(get=GetCommandTimeout,put=PutCommandTimeout)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtimeout_HV10294196.xml; 
 
long GetConnectionTimeout( ); 
void PutConnectionTimeout( long plTimeout ); 
__declspec(property(get=GetConnectionTimeout,put=PutConnectionTimeout)) 
 long Invalid DDUE based on source, error:link not allowed in code, link filename:mdproconnectiontimeout_HV10294222.xml; 
 
_bstr_t GetVersion( ); 
__declspec(property(get=GetVersion)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproversion_HV10294926.xml; 
 
ErrorsPtr GetErrors( ); 
__declspec(property(get=GetErrors)) ErrorsPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolerrors_HV10294338.xml; 
 
_bstr_t GetDefaultDatabase( ); 
void PutDefaultDatabase( _bstr_t pbstr ); 
__declspec(property(get=GetDefaultDatabase,put=PutDefaultDatabase)) 
 _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodefaultdatabase_HV10294288.xml; 
 
enum IsolationLevelEnum GetIsolationLevel( ); 
void PutIsolationLevel( enum IsolationLevelEnum Level ); 
__declspec(property(get=GetIsolationLevel,put=PutIsolationLevel)) enum 
 IsolationLevelEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdproisolationlevel_HV10294459.xml; 
 
long GetAttributes( ); 
void PutAttributes( long plAttr ); 
__declspec(property(get=GetAttributes,put=PutAttributes)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdproattributes_HV10294098.xml; 
 
enum CursorLocationEnum GetCursorLocation( ); 
void PutCursorLocation( enum CursorLocationEnum plCursorLoc ); 
__declspec(property(get=GetCursorLocation,put=PutCursorLocation)) enum 
 CursorLocationEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocursorlocation_HV10294254.xml; 
 
enum ConnectModeEnum GetMode( ); 
void PutMode( enum ConnectModeEnum plMode ); 
__declspec(property(get=GetMode,put=PutMode)) enum ConnectModeEnum 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromode_HV10294518.xml; 
 
_bstr_t GetProvider( ); 
void PutProvider( _bstr_t pbstr ); 
__declspec(property(get=GetProvider,put=PutProvider)) _bstr_t 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprovider_HV10294673.xml; 
 
long GetState( ); 
__declspec(property(get=GetState)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml; 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]