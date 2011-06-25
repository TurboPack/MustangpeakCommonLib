unit MPDataObject;

// Version 2.1.0
//
// The contents of this file are subject to the Mozilla Public License
// Version 1.1 (the "License"); you may not use this file except in compliance
// with the License. You may obtain a copy of the License at http://www.mozilla.org/MPL/
//
// Alternatively, you may redistribute this library, use and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation;
// either version 2.1 of the License, or (at your option) any later version.
// You may obtain a copy of the LGPL at http://www.gnu.org/copyleft/.
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for the
// specific language governing rights and limitations under the License.
//
// The initial developer of this code is Jim Kueneman <jimdk@mindspring.com>
// Special thanks to the following in no particular order for their help/support/code
//    Danijel Malik, Robert Lee, Werner Lehmann, Alexey Torgashin, Milan Vandrovec
//
//----------------------------------------------------------------------------

interface

{.$DEFINE GX_DEBUG_COMMONDATAOBJECT}

{$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  {$DEFINE GX_DEBUG}
{$ENDIF}

{$I Compilers.inc}
{$I Options.inc}
{$I ..\Include\Debug.inc}
{$I ..\Include\Addins.inc}

{$ifdef COMPILER_12_UP}
  {$WARN IMPLICIT_STRING_CAST       OFF}
 {$WARN IMPLICIT_STRING_CAST_LOSS  OFF}
{$endif COMPILER_12_UP}

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  ActiveX,
  ShlObj,
  ShellAPI,
  {$IFDEF GX_DEBUG}
  DbugIntf,
  {$ENDIF}
  MPShellTypes,
  MPCommonUtilities,
  MPCommonObjects,
  {$IFDEF TNTSUPPORT}
  TntStdCtrls,
  TntClasses,
    {$IFDEF COMPILER_10_UP}
    WideStrings,
    {$ELSE}
    TntWideStrings,
    {$ENDIF}
  {$ENDIF}
  AxCtrls;

const
  // Standard Shell Formats
  CFSTR_LOGICALPERFORMEDDROPEFFECT = 'Logical Performed DropEffect';
  CFSTR_PREFERREDDROPEFFECT = 'Preferred DropEffect';
  CFSTR_PERFORMEDDROPEFFECT = 'Performed DropEffect';
  CFSTR_PASTESUCCEEDED = 'Paste Succeeded';
  CFSTR_INDRAGLOOP = 'InShellDragLoop';
  CFSTR_SHELLIDLISTOFFSET = 'Shell Object Offsets';
  SIZE_SHELLDRAGLOOPDATA = 4;

type
  PPerformedDropEffect = ^TPerformedDropEffect;
  TPerformedDropEffect = (
    effectNone,    // No Operation (DROPEFFECT_NONE)
    effectCopy,    // Operation was a copy (DROPEFFECT_COPY)
    effectMove,    // Operation was a move (DROPEFFECT_MOVE)
    effectLink     // Operation was a link (DROPEFFECT_LINK)
  );

type
  TFormatEtcArray = array of TFormatEtc;
  TDataObjectInfo = record
    FormatEtc: TFormatEtc;
    StgMedium: TStgMedium;
    OwnedByDataObject: Boolean
  end;
  TDataObjectInfoArray = array of TDataObjectInfo;

type
  // This is just a WAG for the size.  I can't imagine it ever needing this
  // many formats
  TeltArray = array[0..255] of TFormatEtc;

  {$IFNDEF COMPILER_6_UP}
  PCardinal = ^Cardinal;
  {$ENDIF}
  
//------------------------------------------------------------------------------
// TCommonEnumFormatEtc :
//       Implements the IEnumFormatEtc interface for the IDataObject
// implementation.  This interface is called by a potential droptarget to see it
// the IDataObject contains data that the target knows how to handle and would
// like a shot at accepting it.
//-------------------------------------------------------------------------------

  TCommonEnumFormatEtc = class(TInterfacedObject, IEnumFormatEtc)
  private
    FInternalIndex: integer;
    FFormats: TFormatEtcArray;
  protected
    { IEnumFormatEtc }
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out Enum: IEnumFormatEtc): HResult; stdcall;

    property InternalIndex: integer read FInternalIndex write FInternalIndex;
  public
    constructor Create;
    destructor Destroy; override;

    procedure SetFormatLength(Size: Integer);
    property Formats: TFormatEtcArray read FFormats write FFormats;
  end;

//-------------------------------------------------------------------------------}
// TCommonDataObject :                                                                 }
//       Implements the IDataObject interface.  This interface is called by a    }
// potential droptarget to see it the IDataObject contains data that the target  }
// knows how to handle and would like a shot at accepting it.                    }
//-------------------------------------------------------------------------------}

  ICommonDataObject = interface(IDataObject)
    ['{F8B3EE47-C6C1-4FE3-9D94-757AA35DC038}']
    function AssignDragImage(Image: TBitmap; HotSpot: TPoint; TransparentColor: TColor): Boolean;
    function SaveGlobalBlock(Format: TClipFormat; MemoryBlock: Pointer; MemoryBlockSize: integer): Boolean;
    function LoadGlobalBlock(Format: TClipFormat; var MemoryBlock: Pointer): Boolean;
  end;

  ICommonDataObject2 = interface(ICommonDataObject)
  ['{1E707CE8-F2E7-4566-A87E-630859509F97}']
    function MultiFolder: Boolean;      // Tests to see if the DataObject a multiFolder dataobject
    function MultiFolderVerified: Boolean;    // Tests to see if the context of the data object are ordered and correctly filtered to not give problems to shell functions
    procedure SetMultiFolder(IsSet: Boolean);
    procedure SetMultiFolderVerified(IsVerified: Boolean);
  end;

  TGetDataEvent = procedure(Sender: TObject; const FormatEtcIn: TFormatEtc; var Medium: TStgMedium; var Handled: Boolean) of object;
  TQueryGetDataEvent = procedure(Sender: TObject; const FormatEtcIn: TFormatEtc; var FormatAvailable: Boolean; var Handled: Boolean) of object;
         
  TCommonDataObject = class(TObject, IUnknown, IDataObject, ICommonDataObject, ICommonExtractObj, ICommonDataObject2)
  private
    FIsMultiFolderVerified: Boolean;
    FIsMultiFolder: Boolean;
    FNotifyThreadRef: Boolean;
    FRefCount: Integer;
    FReferenceCounted: Boolean;
  protected
    FAdviseHolder: IDataAdviseHolder;
    FFormats: TDataObjectInfoArray;
    FOnGetData: TGetDataEvent;
    FOnQueryGetData: TQueryGetDataEvent;  // Reference to an OLE supplied implementation for advising.

    // IUnknown
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    // IDataObject
    function CanonicalIUnknown(TestUnknown: IUnknown): IUnknown;
    function DAdvise(const formatetc: TFormatEtc; advf: Longint; const advSink: IAdviseSink; out dwConnection: Longint): HResult; virtual; stdcall;
    function DUnadvise(dwConnection: Longint): HResult; virtual; stdcall;
    function EnumDAdvise(out enumAdvise: IEnumStatData): HResult;virtual; stdcall;
    function EnumFormatEtc(dwDirection: Longint; out enumFormatEtc: IEnumFormatEtc): HResult;virtual; stdcall;
    function EqualFormatEtc(FormatEtc1, FormatEtc2: TFormatEtc): Boolean;
    function FindFormatEtc(TestFormatEtc: TFormatEtc): integer;
    function GetCanonicalFormatEtc(const formatetc: TFormatEtc; out formatetcOut: TFormatEtc): HResult;virtual; stdcall;
    function GetData(const FormatEtcIn: TFormatEtc; out Medium: TStgMedium): HResult;virtual; stdcall;
    function GetDataHere(const formatetc: TFormatEtc; out medium: TStgMedium): HResult;virtual; stdcall;
    function HGlobalClone(HGlobal: THandle): THandle;
    function QueryGetData(const formatetc: TFormatEtc): HResult;virtual; stdcall;
    function SetData(const formatetc: TFormatEtc; var medium: TStgMedium; fRelease: BOOL): HResult;virtual; stdcall;
    procedure DoGetCustomFormats(dwDirection: Integer; var Formats: TFormatEtcArray); virtual;
    procedure DoOnGetData(const FormatEtcIn: TFormatEtc; var Medium: TStgMedium; var Handled: Boolean); virtual;
    procedure DoOnQueryGetData(const FormatEtcIn: TFormatEtc; var FormatAvailable: Boolean; var Handled: Boolean); virtual;
    function RetrieveOwnedStgMedium(Format: TFormatEtc; var StgMedium: TStgMedium): HRESULT;
    function StgMediumIncRef(const InStgMedium: TStgMedium; var OutStgMedium: TStgMedium; CopyInMedium: Boolean): HRESULT;

    // ICommonExtractObj
    function GetObj: TObject;

    // ICommonDataObject2
    function MultiFolder: Boolean;
    function MultiFolderVerified: Boolean;
    procedure SetMultiFolder(IsSet: Boolean);
    procedure SetMultiFolderVerified(IsVerified: Boolean);

    property AdviseHolder: IDataAdviseHolder read FAdviseHolder;
    property Formats: TDataObjectInfoArray read FFormats write FFormats;
    property NotifyThreadRef: Boolean read FNotifyThreadRef write FNotifyThreadRef;
    property Obj: TObject read GetObj;
    property RefCount: Integer read FRefCount;
  public
    constructor Create(IsReferenceCounted: Boolean = True);
    destructor Destroy; override;
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
    class function NewInstance: TObject; override;
    function AssignDragImage(Image: TBitmap; HotSpot: TPoint; TransparentColor: TColor): Boolean;
    function GetUserData(Format: TFormatEtc; var StgMedium: TStgMedium): Boolean; virtual;
    function LoadGlobalBlock(Format: TClipFormat; var MemoryBlock: Pointer): Boolean;
    function SaveGlobalBlock(Format: TClipFormat; MemoryBlock: Pointer; MemoryBlockSize: integer): Boolean;

    property IsMultiFolderVerified: Boolean read FIsMultiFolderVerified write FIsMultiFolderVerified;
    property IsMultiFolder: Boolean read FIsMultiFolder write FIsMultiFolder;
    property ReferenceCounted: Boolean read FReferenceCounted write FReferenceCounted;
    property OnGetData: TGetDataEvent read FOnGetData write FOnGetData;
    property OnQueryGetData: TQueryGetDataEvent read FOnQueryGetData write FOnQueryGetData;
  end;
  TCommonDataObjectClass = class of TCommonDataObject;

  TCommonClipboardFormat = class
  public
    function DataObjectContainsFormat(DataObject: IDataObject): Boolean;
    function GetFormatEtc: TFormatEtc; virtual;
    function LoadFromClipboard: Boolean; virtual;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; virtual; abstract;
    function SaveToClipboard: Boolean; virtual;
    function SaveToDataObject(DataObject: IDataObject): Boolean; virtual; abstract;
  end;

  TCommonStreamClipFormat = class(TCommonClipboardFormat)
  public
    function GetFormatEtc: TFormatEtc; override;
    function LoadFromClipboard: Boolean; override;
    function LoadFromDataObject(DataObject: IDataObject; CoolStream: TCommonStream): Boolean; reintroduce;
    function SaveToClipboard: Boolean; override;
    function SaveToDataObject(DataObject: IDataObject; CoolStream: TCommonStream): Boolean; reintroduce;
  end;

// Simpifies dealing with the CFSTR_FILEGROUPDESCRIPTOR format
  TDescriptorAArray = array of TFileDescriptorA;
  TDescriptorWArray = array of TFileDescriptorW;

  TFileGroupDescriptorA = class(TCommonClipboardFormat)
  private
    FStream: TStream;
    function GetDescriptorCount: Integer;
    function GetFileDescriptorA(Index: Integer): TFileDescriptorA;
    procedure SetFileDescriptor(Index: Integer;
      const Value: TFileDescriptorA);
  protected
    FFileDescriptors: TDescriptorAArray;
    property Stream: TStream read FStream write FStream;
  public
    destructor Destroy; override;
    procedure AddFileDescriptor(FileDescriptor: TFileDescriptorA);
    procedure DeleteFileDescriptor(Index: integer);
    function GetFormatEtc: TFormatEtc; override;
    function FillDescriptor(FileName: AnsiString): TFileDescriptorA;
    function GetFileStream(const DataObject: IDataObject; FileIndex: Integer): TStream;
    procedure LoadFileGroupDestriptor(FileGroupDiscriptor: PFileGroupDescriptorA);
    function LoadFromClipboard: Boolean; override;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; override;
    function SaveToClipboard: Boolean; override;
    function SaveToDataObject(DataObject: IDataObject): Boolean; override;

    property DescriptorCount: Integer read GetDescriptorCount;
    property FileDescriptor[Index: Integer]: TFileDescriptorA read GetFileDescriptorA write SetFileDescriptor;
  end;

  TFileGroupDescriptorW = class(TCommonClipboardFormat)
  private
    FStream: TStream;
    function GetDescriptorCount: Integer;
    function GetFileDescriptorW(Index: Integer): TFileDescriptorW;
    procedure SetFileDescriptor(Index: Integer;
      const Value: TFileDescriptorW);
  protected
    FFileDescriptors: TDescriptorWArray;
    property Stream: TStream read FStream write FStream;
  public
    destructor Destroy; override;
    procedure AddFileDescriptor(FileDescriptor: TFileDescriptorW);
    procedure DeleteFileDescriptor(Index: integer);
    function FillDescriptor(FileName: WideString): TFileDescriptorW;
    function GetFileStream(const DataObject: IDataObject; FileIndex: Integer): TStream;
    function GetFormatEtc: TFormatEtc; override;
    procedure LoadFileGroupDestriptor(FileGroupDiscriptor: PFileGroupDescriptorW);
    function LoadFromClipboard: Boolean; override;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; override;
    function SaveToClipboard: Boolean; override;
    function SaveToDataObject(DataObject: IDataObject): Boolean; override;

    property DescriptorCount: Integer read GetDescriptorCount;
    property FileDescriptor[Index: Integer]: TFileDescriptorW read GetFileDescriptorW write SetFileDescriptor;
  end;

  // Simpifies dealing with the CF_HDROP format
  TCommonHDrop = class(TCommonClipboardFormat)
  private
    procedure SetDropFiles(const Value: PDropFiles);
    function GetHDropStruct: THandle;
  protected
    FDropFiles: PDropFiles;
    FStructureSize: integer;
    FFileCount: integer;

    procedure AllocStructure(Size: integer);
    function CalculateDropFileStructureSizeA(Value: PDropFiles): integer;
    function CalculateDropFileStructureSizeW(Value: PDropFiles): integer;
    function FileCountA: Integer;
    function FileCountW: Integer;
    function FileNameA(Index: integer): AnsiString;
    function FileNameW(Index: integer): WideString;
    procedure FreeStructure; // Frees memory allocated
  public
    destructor Destroy; override;
    function AssignFromClipboard: Boolean;
    function LoadFromClipboard: Boolean; override;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; override;
    function FileCount: integer;
    function FileName(Index: integer): WideString;
    function GetFormatEtc: TFormatEtc; override;
    procedure AssignFilesA(FileList: TStringList);
    {$IFDEF TNTSUPPORT}
    procedure AssignFilesW(FileList: TWideStrings);
    procedure FileNamesW(FileList: TWideStrings);
    {$ENDIF}
    {$IFDEF  COMPILER_12_UP}
    procedure AssignFilesW(FileList: TStrings);
    procedure FileNamesW(FileList: TStrings);
    {$ENDIF}
    procedure FileNamesA(FileList: TStrings);

    property HDropStruct: THandle read GetHDropStruct;
    function SaveToClipboard: Boolean; override;
    function SaveToDataObject(DataObject: IDataObject): Boolean; override;
    property StructureSize: integer read FStructureSize;
    property DropFiles: PDropFiles read FDropFiles write SetDropFiles;
  end;

  // Simpifies dealing with the CFSTR_SHELLIDLIST format
type
  TCommonShellIDList = class(TCommonClipboardFormat)
  private
    FCIDA: PIDA;
    function GetCIDASize: integer;
    function InternalChildPIDL(Index: integer): PItemIDList;
    function InternalParentPIDL: PItemIDList;
    procedure SetCIDA(const Value: PIDA);
  public
    function AbsolutePIDL(Index: integer): PItemIDList;
    procedure AbsolutePIDLs(APIDLList: TCommonPIDLList);
    procedure AssignPIDLs(APIDLList: TCommonPIDLList);
    destructor Destroy; override;
    function GetFormatEtc: TFormatEtc; override;
    function LoadFromClipboard: Boolean; override;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; override;
    function ParentPIDL: PItemIDList;
    function PIDLCount: integer;
    function RelativePIDL(Index: integer): PItemIDList;
    procedure RelativePIDLs(APIDLList: TCommonPIDLList);
    function SaveToClipboard: Boolean; override;
    function SaveToDataObject(DataObject: IDataObject): Boolean; override;

    property CIDA: PIDA read FCIDA write SetCIDA;
    property CIDASize: integer read GetCIDASize;
  end;

  // Simpilies dealing with the CFSTR_LOGICALPERFORMEDDROPEFFECT format
  TCommonLogicalPerformedDropEffect = class(TCommonClipboardFormat)
  private
    FAction: TPerformedDropEffect;
  public
    function GetFormatEtc: TFormatEtc; override;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; override;
    function SaveToDataObject(DataObject: IDataObject): Boolean; override;

    property Action: TPerformedDropEffect read FAction write FAction;
  end;

  // Simpilies dealing with the CFSTR_PerferredDropEffect format
  TCommonPreferredDropEffect = class(TCommonLogicalPerformedDropEffect)
  public
    function GetFormatEtc: TFormatEtc; override;
  end;

  TCommonInShellDragLoop = class(TCommonClipboardFormat)
  private
    FData: Cardinal;
  public
    function GetFormatEtc: TFormatEtc; override;
    function LoadFromDataObject(DataObject: IDataObject): Boolean; override;
    function SaveToDataObject(DataObject: IDataObject): Boolean; override;
    property Data: Cardinal read FData write FData;
  end;

  function FillFormatEtc(cfFormat: Word; ptd: PDVTargetDevice = nil;
    dwAspect: Longint = DVASPECT_CONTENT; lindex: Longint = -1; tymed: Longint = TYMED_HGLOBAL): TFormatEtc;
  function DataObjectSupportsShell(const DataObj: IDataObject): Boolean;
  function DataObjectContainsPIDL(APIDL: PItemIDList; const DataObj: IDataObject): Boolean;

  function HDropFormat: TFormatEtc;
  function ShellIDListFormat: TFormatEtc;
  function FileDescriptorAFormat: TFormatEtc;
  function FileDescriptorWFormat: TFormatEtc;

var
  CF_SHELLIDLIST,
  CF_PERFORMEDDROPEFFECT,
  CF_PASTESUCCEEDED,
  CF_INDRAGLOOP,
  CF_SHELLIDLISTOFFSET,
  CF_LOGICALPERFORMEDDROPEFFECT,
  CF_PREFERREDDROPEFFECT,
  CF_FILECONTENTS,
  CF_FILEDESCRIPTORA,
  CF_FILEDESCRIPTORW: TClipFormat;

implementation

uses
  MPShellUtilities, Math;

var
  PIDLMgr: TCommonPIDLManager;
  ShellILIsEqual: function(PIDL1: PItemIDList; PIDL2: PItemIDList): LongBool; stdcall;
  ShellILIsEqualChecked: Boolean = False;


function LoadShellILIsEqual: Boolean;
begin
  if not ShellILIsEqualChecked then
  begin
    ShellILIsEqual := GetProcAddress(GetModuleHandleA(PAnsiChar( AnsiString(Shell32))), PAnsiChar(21));
    ShellILIsEqualChecked := True;
  end;
  Result := Assigned(ShellILIsEqual)
end;

function DataObjectContainsPIDL(APIDL: PItemIDList; const DataObj: IDataObject): Boolean;
var
  ShellIDList: TCommonShellIDList;
  i: Integer;
  P: PItemIDList;
begin
  Result := False;
  ;
  if Assigned(DataObj) and Assigned(APIDL) and LoadShellILIsEqual then
  begin
    if Succeeded(DataObj.QueryGetData(ShellIDListFormat)) then
    begin
      ShellIDList := TCommonShellIDList.Create;
      try
        if ShellIDList.LoadFromDataObject(DataObj) then
        begin
          i := 0;
          while not Result and (i < ShellIDList.PIDLCount) do
          begin
            P := ShellIDList.AbsolutePIDL(i);
            Result := ShellILIsEqual(P, APIDL);
            PIDLMgr.FreePIDL(P);
            Inc(i)
          end
        end
      finally
        ShellIDList.Free
      end
    end
  end
end;

function DataObjectSupportsShell(const DataObj: IDataObject): Boolean;
begin
  Result := False;
  if Assigned(DataObj) then
  begin
    Result := Succeeded(DataObj.QueryGetData(HDropFormat)) or
      Succeeded(DataObj.QueryGetData(ShellIDListFormat)) or
      Succeeded(DataObj.QueryGetData(FileDescriptorAFormat)) or
      Succeeded(DataObj.QueryGetData(FileDescriptorWFormat))
  end
end;

function HDropFormat: TFormatEtc;
begin
  Result.cfFormat := CF_HDROP; // This guy is always registered for all applications
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL
end;

function ShellIDListFormat: TFormatEtc;
begin
  Result.cfFormat := CF_SHELLIDLIST;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL;
end;

function FileDescriptorAFormat: TFormatEtc; 
begin 
  Result.cfFormat := CF_FILEDESCRIPTORA; 
  Result.ptd := nil; 
  Result.dwAspect := DVASPECT_CONTENT; 
  Result.lindex := -1; 
  Result.tymed := TYMED_HGLOBAL 
end; 

function FileDescriptorWFormat: TFormatEtc; 
begin 
  Result.cfFormat := CF_FILEDESCRIPTORW; 
  Result.ptd := nil; 
  Result.dwAspect := DVASPECT_CONTENT; 
  Result.lindex := -1; 
  Result.tymed := TYMED_HGLOBAL 
end;

function FillFormatEtc(cfFormat: Word; ptd: PDVTargetDevice = nil;
  dwAspect: Longint = DVASPECT_CONTENT; lindex: Longint = -1; tymed: Longint = TYMED_HGLOBAL): TFormatEtc;
begin
  Result.cfFormat := cfFormat;
  Result.ptd := ptd;
  Result.dwAspect := dwAspect;
  Result.lindex := lindex;
  Result.tymed := tymed
end;

{ TCommonClipboardFormat }
function TCommonClipboardFormat.DataObjectContainsFormat(DataObject: IDataObject): Boolean;
begin
  Result := DataObject.QueryGetData(GetFormatEtc) = S_OK
end;


function TCommonClipboardFormat.GetFormatEtc: TFormatEtc;
begin
  FillChar(Result, SizeOf(Result), #0);
end;

function TCommonClipboardFormat.LoadFromClipboard: Boolean;
begin
  Result := False;
end;

function TCommonClipboardFormat.SaveToClipboard: Boolean;
begin
  Result := False;
end;

{ TLogicalPerformedDropEffect }

function TCommonLogicalPerformedDropEffect.GetFormatEtc: TFormatEtc;
begin
  Result.cfFormat := CF_LOGICALPERFORMEDDROPEFFECT;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL
end;

function TCommonLogicalPerformedDropEffect.LoadFromDataObject(DataObject: IDataObject): Boolean;
var
  Ptr: PPerformedDropEffect;
  StgMedium: TStgMedium;
begin
  Result := False;
  FillChar(StgMedium, SizeOf(StgMedium), #0);

  if Succeeded(DataObject.GetData(GetFormatEtc, StgMedium)) then
  try
    Ptr := GlobalLock(StgMedium.hGlobal);
    try
      if Assigned(Ptr) then
      begin
        FAction := Ptr^;
        Result := True;
      end;
    finally
      GlobalUnLock(StgMedium.hGlobal);
    end
  finally
    ReleaseStgMedium(StgMedium)
  end
end;

function TCommonLogicalPerformedDropEffect.SaveToDataObject(DataObject: IDataObject): Boolean;
var
  Ptr: PPerformedDropEffect;
  StgMedium: TStgMedium;
begin
  FillChar(StgMedium, SizeOf(StgMedium), #0);

  StgMedium.hGlobal := GlobalAlloc(GPTR, SizeOf(FAction));
  Ptr := GlobalLock(StgMedium.hGlobal);
  try
    Ptr^ := FAction;
    StgMedium.tymed := TYMED_HGLOBAL;
    Result := Succeeded(DataObject.SetData(GetFormatEtc, StgMedium, True))
  finally
    GlobalUnLock(StgMedium.hGlobal);
  end
end;

{ THDrop }

procedure TCommonHDrop.AllocStructure(Size: integer);
begin
  FreeStructure;
  GetMem(FDropFiles, Size);
  FStructureSize := Size;
  FillChar(FDropFiles^, Size, #0);
end;

procedure TCommonHDrop.AssignFilesA(FileList: TStringList);
var
  i: Integer;
  Size: integer;
  Path: PAnsiChar;
begin
  if Assigned(FileList) then
  begin
    FreeStructure;
    Size := 0;
    for i := 0 to FileList.Count - 1 do
      Inc(Size, Length(FileList[i]) + SizeOf(AnsiChar)); // add spot for the null
    Inc(Size, SizeOf(TDropFiles));
    Inc(Size, SizeOf(AnsiChar)); // room for the terminating null
    AllocStructure(Size);
    DropFiles.pFiles := SizeOf(TDropFiles);
    DropFiles.pt.x := 0;
    DropFiles.pt.y := 0;
    DropFiles.fNC := False;
    DropFiles.fWide := False;  // Don't support wide char let NT convert it
    Path := PAnsiChar(FDropFiles) + FDropFiles.pFiles;
    for i := 0 to FileList.Count - 1 do
    begin
      MoveMemory(Path, Pointer(AnsiString( FileList[i])), Length(FileList[i]));
      Inc(Path, Length(FileList[i]) + 1); // skip over the single null #0
    end
  end
end;

function TCommonHDrop.AssignFromClipboard: Boolean;
var
  Handle: THandle;
  Ptr: PDropFiles;
begin
  Result := False;
  Handle := 0;
  OpenClipboard(Application.Handle);
  try
    Handle := GetClipboardData(CF_HDROP);
    if Handle <> 0 then
    begin
      Ptr := GlobalLock(Handle);
      if Assigned(Ptr) then
      begin
        DropFiles := Ptr;
        Result := True;
      end;
    end;
  finally
    CloseClipboard;
    GlobalUnLock(Handle);
  end;
end;

function TCommonHDrop.CalculateDropFileStructureSizeA(
  Value: PDropFiles): integer;
var
  Head: PAnsiChar;
  Len: integer;
begin
  if Assigned(Value) then
  begin
    Result := Value^.pFiles;
    Head := PAnsiChar( Value) + Value^.pFiles;
    Len := lstrlenA(Head);
    while Len > 0 do
    begin
      Result := Result + Len + 1;
      Head := Head + Len + 1;
       Len := lstrlenA(Head);
    end;
    Inc(Result, 1); // Add second null
  end else
    Result := 0
end;

function TCommonHDrop.CalculateDropFileStructureSizeW(
  Value: PDropFiles): integer;
var
  Head: PAnsiChar;
  Len: integer;
begin
  if Assigned(Value) then
  begin
    Result := Value^.pFiles;
    Head := PAnsiChar( Value) + Value^.pFiles;
    Len := 2 * (lstrlenW(PWideChar( Head)));
    while Len > 0 do
    begin
      Result := Result + Len + 2;
      Head := Head + Len + 2;
       Len := 2 * (lstrlenW(PWideChar( Head)));
    end;
    Inc(Result, 2); // Add second null
  end else
    Result := 0
end;

destructor TCommonHDrop.Destroy;
begin
  FreeStructure;
  inherited;
end;

function TCommonHDrop.FileCount: integer;
begin
  if Assigned(DropFiles) then
  begin
    if FFileCount = 0 then
    begin
      if DropFiles.fWide then
        Result := FileCountW
      else
        Result := FileCountA;
       FFileCount := Result;
    end;
  end else
   FFileCount := 0;
  Result := FFileCount
end;

function TCommonHDrop.FileCountA: Integer;
var
  Head: PAnsiChar;
  Len: integer;
begin
  Result := 0;
  if Assigned(DropFiles) then
  begin
    Head := PAnsiChar( DropFiles) + DropFiles^.pFiles;
    Len := lstrlenA(Head);
    while Len > 0 do
    begin
      Head := Head + Len + 1;
      Inc(Result);
      Len := lstrlenA(Head);
    end
  end
end;

function TCommonHDrop.FileCountW: Integer;
var
  Head: PAnsiChar;
  Len: integer;
begin
  Result := 0;
  if Assigned(DropFiles) then
  begin
    Head := PAnsiChar( DropFiles) + DropFiles^.pFiles;
    Len := 2 * (lstrlenW(PWideChar( Head)));
    while Len > 0 do
    begin
      Head := Head + Len + 2;
      Inc(Result);
      Len := 2 * (lstrlenW(PWideChar( Head)));
    end
  end;
end;

function TCommonHDrop.FileName(Index: integer): WideString;
begin
  if Assigned(DropFiles) then
  begin
    if DropFiles.fWide then
      Result := FileNameW(Index)
    else
      Result := FileNameA(Index)
  end
end;

function TCommonHDrop.FileNameA(Index: integer): AnsiString;
var
  Head: PAnsiChar;
  PathNameCount: integer;
  Done: Boolean;
  Len: integer;
begin
  PathNameCount := 0;
  Done := False;
  if Assigned(DropFiles) then
  begin
    Head := PAnsiChar( DropFiles) + DropFiles^.pFiles;
    Len := lstrlenA(Head);
    while (not Done) and (PathNameCount < FileCount) do
    begin
      if PathNameCount = Index then
      begin
        SetLength(Result, Len + 1);
        CopyMemory(@Result[1], Head, Len + 1); // Include the NULL
        Done := True;
      end;
      Head := Head + Len + 1;
      Inc(PathNameCount);
      Len := lstrlenA(Head);
    end
  end
end;

{$IFDEF TNTSUPPORT}
procedure TCommonHDrop.AssignFilesW(FileList: TWideStrings);
var
  i: Integer;
  Size: integer;
  Path: PAnsiChar;
  ByteSize: Integer;
begin
  if Assigned(FileList) then
  begin
    FreeStructure;
    Size := 0;
    if UnicodeStringLists then
      ByteSize := 2
    else
      ByteSize := 1;
    for i := 0 to FileList.Count - 1 do
      Inc(Size, (Length(FileList[i])+1)*(SizeOf(AnsiChar)*ByteSize)); // add spot for the null
    Inc(Size, SizeOf(TDropFiles));
    Inc(Size, SizeOf(AnsiChar)*2); // room for the terminating null
    AllocStructure(Size);
    DropFiles.pFiles := SizeOf(TDropFiles);
    DropFiles.pt.x := 0;
    DropFiles.pt.y := 0;
    DropFiles.fNC := False;
    if UnicodeStringLists then
      Integer( DropFiles.fWide ):= 1
    else
      Integer( DropFiles.fWide) := 0;
    Path := PAnsiChar(FDropFiles) + FDropFiles.pFiles;
    for i := 0 to FileList.Count - 1 do
    begin
      MoveMemory(Path, Pointer(FileList[i]), Length(FileList[i])*ByteSize);
      Inc(Path, (Length(FileList[i]) + 1)*ByteSize); // skip over the single null #0
    end
  end
end;

procedure TCommonHDrop.FileNamesW(FileList: TWideStrings);
var
  i: integer;
begin
  if Assigned(FileList) then
  begin
    for i := 0 to FileCount - 1 do
      FileList.Add(FileNameW(i));
  end;
end;
{$ENDIF}

{$IFDEF COMPILER_12_UP}
procedure TCommonHDrop.AssignFilesW(FileList: TStrings);
var
  i: Integer;
  Size: integer;
  Path: PAnsiChar;
  ByteSize: Integer;
begin
  if Assigned(FileList) then
  begin
    FreeStructure;
    Size := 0;
    if UnicodeStringLists then
      ByteSize := 2
    else
      ByteSize := 1;
    for i := 0 to FileList.Count - 1 do
      Inc(Size, (Length(FileList[i])+1)*(SizeOf(AnsiChar)*ByteSize)); // add spot for the null
    Inc(Size, SizeOf(TDropFiles));
    Inc(Size, SizeOf(AnsiChar)*2); // room for the terminating null
    AllocStructure(Size);
    DropFiles.pFiles := SizeOf(TDropFiles);
    DropFiles.pt.x := 0;
    DropFiles.pt.y := 0;
    DropFiles.fNC := False;
    if UnicodeStringLists then
      Integer( DropFiles.fWide ):= 1
    else
      Integer( DropFiles.fWide) := 0;
    Path := PAnsiChar(FDropFiles) + FDropFiles.pFiles;
    for i := 0 to FileList.Count - 1 do
    begin
      MoveMemory(Path, Pointer(FileList[i]), Length(FileList[i])*ByteSize);
      Inc(Path, (Length(FileList[i]) + 1)*ByteSize); // skip over the single null #0
    end
  end
end;

procedure TCommonHDrop.FileNamesW(FileList: TStrings);
var
  i: integer;
begin
  if Assigned(FileList) then
  begin
    for i := 0 to FileCount - 1 do
      FileList.Add(FileNameW(i));
  end;
end;
{$ENDIF}


procedure TCommonHDrop.FileNamesA(FileList: TStrings);
var
  i: integer;
begin
  if Assigned(FileList) then
  begin
    for i := 0 to FileCount - 1 do
      FileList.Add(FileName(i));
  end;
end;

function TCommonHDrop.FileNameW(Index: integer): WideString;
var
  Head: PAnsiChar;
  PathNameCount: integer;
  Done: Boolean;
  Len: integer;
begin
  PathNameCount := 0;
  Done := False;
  if Assigned(DropFiles) then
  begin
    Head := PAnsiChar( DropFiles) + DropFiles^.pFiles;
    Len := 2 * (lstrlenW(PWideChar( Head)));
    while (not Done) and (PathNameCount < FileCount) do
    begin
      if PathNameCount = Index then
      begin
        SetLength(Result, (Len + 1) div 2);
        CopyMemory(@Result[1], Head, Len + 2); // Include the NULL
        Done := True;
      end;
      Head := Head + Len + 2;
      Inc(PathNameCount);
      Len := 2 * (lstrlenW(PWideChar( Head)));
    end
  end
end;

procedure TCommonHDrop.FreeStructure;
begin
  FFileCount := 0;
  if Assigned(FDropFiles) then
    FreeMem(FDropFiles, FStructureSize);
  FDropFiles := nil;
  FStructureSize := 0;
end;

function TCommonHDrop.GetFormatEtc: TFormatEtc;
begin
  Result.cfFormat := CF_HDROP; // This guy is always registered for all applications
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL
end;

function TCommonHDrop.GetHDropStruct: THandle;
var
  Files: PDropFiles;
begin
  Result := GlobalAlloc(GHND, StructureSize);
  Files := GlobalLock(Result);
  try
    MoveMemory(Files, FDropFiles, StructureSize);
  finally
    GlobalUnlock(Result)
  end;
end;

function TCommonHDrop.LoadFromClipboard: Boolean;
var
  Handle: THandle;
  Ptr: PDropFiles;
begin
  Result := False;
  Handle := 0;
  OpenClipboard(Application.Handle);
  try
    Handle := GetClipboardData(CF_HDROP);
    if Handle <> 0 then
    begin
      Ptr := GlobalLock(Handle);
      if Assigned(Ptr) then
      begin
        DropFiles := Ptr;
        Result := True;
      end;
    end;
  finally
    CloseClipboard;
    GlobalUnLock(Handle);
  end;
end;

function TCommonHDrop.LoadFromDataObject(DataObject: IDataObject): Boolean;
var
  Medium: TStgMedium;
  Files: PDropFiles;
begin
  Result := False;
  if Assigned(DataObject) then
  begin
    if Succeeded(DataObject.GetData(GetFormatEtc, Medium)) then
    try
      Files := GlobalLock(Medium.hGlobal);
      try
        DropFiles := Files
      finally
        GlobalUnlock(Medium.hGlobal)
      end
    finally
      ReleaseStgMedium(Medium)
    end;
    Result := Assigned(DropFiles)
  end
end;

function TCommonHDrop.SaveToClipboard: Boolean;
begin
  Result := False;
  OpenClipboard(Application.Handle);
  try
    SetClipboardData(CF_HDROP, HDropStruct)
  finally
    CloseClipboard;
  end;
end;

function TCommonHDrop.SaveToDataObject(DataObject: IDataObject): Boolean;
var
  Medium: TStgMedium;
begin
  Result := False;
  FillChar(Medium, SizeOf(Medium), #0);
  Medium.tymed := TYMED_HGLOBAL;
  Medium.hGlobal := HDropStruct;
  // Give the block to the DataObject
  if Succeeded(DataObject.SetData(GetFormatEtc, Medium, True)) then
    Result := True
  else
    GlobalFree(Medium.hGlobal)
end;

procedure TCommonHDrop.SetDropFiles(const Value: PDropFiles);
begin
  FreeStructure;
  if Assigned(Value) then
  begin
    if Value.fWide then
      FStructureSize := CalculateDropFileStructureSizeW(Value)
    else
      FStructureSize := CalculateDropFileStructureSizeA(Value);
    AllocStructure(StructureSize);
    CopyMemory(FDropFiles, Value, StructureSize);
  end;
end;

{ TPreferredDropEffect }

function TCommonPreferredDropEffect.GetFormatEtc: TFormatEtc;
begin
  Result.cfFormat := CF_PREFERREDDROPEFFECT;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL
end;

{ TCoolStreamClipFormat }

function TCommonStreamClipFormat.GetFormatEtc: TFormatEtc;
begin
  Result.cfFormat := 0;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL
end;

function TCommonStreamClipFormat.LoadFromClipboard: Boolean;
begin
  Result := False;
  Assert(True=False, 'TCoolStream.LoadFromClipboard not Implemented');
end;

function TCommonStreamClipFormat.LoadFromDataObject(DataObject: IDataObject;
  CoolStream: TCommonStream): Boolean;
var
  Medium: TStgMedium;
  PMem, PSize: Pointer;
  Size: Int64;
begin
  Result := False;
  if Succeeded(DataObject.GetData(GetFormatEtc, Medium)) then
  begin
    PMem := GlobalLock(Medium.hGlobal);
    try
      PSize := @Size;
      // Get the size of the Stream
      MoveMemory(PMem, PSize, SizeOf(Int64));

      CoolStream.Seek(0, soFromBeginning);
      CoolStream.SetSize(Size);

      Inc(PAnsiChar( PMem), SizeOf(Int64));

      // Copy the data to the stream
      MoveMemory(PMem, CoolStream.Memory, Size);
    finally
      GlobalUnlock(Medium.hGlobal);
      ReleaseStgMedium(Medium);
    end
  end
end;

function TCommonStreamClipFormat.SaveToClipboard: Boolean;
begin
  Result := False;
  Assert(True=False, 'TCoolStream.SaveToClipboard not Implemented');
end;

function TCommonStreamClipFormat.SaveToDataObject(DataObject: IDataObject;
  CoolStream: TCommonStream): Boolean;
var
  Medium: TStgMedium;
  Mem: HGlobal;
  PMem: Pointer;
  Size: Int64;
  PSize: Pointer;
begin
  Mem := GlobalAlloc(GMEM_SHARE  or GMEM_MOVEABLE, CoolStream.Size + SizeOf(Int64));
  PMem := GlobalLock(Mem);
  try
    FillChar(Medium, SizeOf(Medium), #0);
    
    PSize := @Size;
    
    // Write the size of the Stream to the block
    Size := CoolStream.Size;
    MoveMemory(PMem, PSize, SizeOf(Int64));

    // Copy the Stream to a global memory block
    MoveMemory(PMem, CoolStream.Memory, CoolStream.Size);

    // Transfer the data with a global memory block
    Medium.tymed := TYMED_HGLOBAL;
    Medium.hGlobal := Mem;

    // Give it to the DataObject so it may destroy it when the clients are done
    Result := Succeeded(DataObject.SetData(GetFormatEtc, Medium, True))
  finally
    GlobalUnlock(Mem)
  end
end;

{ TShellIDList }

function TCommonShellIDList.AbsolutePIDL(Index: integer): PItemIDList;
{ Appends the single ItemID with the Parent folder to create an Absolute PIDL }
begin
  if Assigned(FCIDA) then
  begin
    Result := PIDLMgr.AppendPIDL(InternalParentPIDL, InternalChildPIDL(Index));
  end else
    Result := nil
end;

procedure TCommonShellIDList.AbsolutePIDLs(APIDLList: TCommonPIDLList);
var
  i: integer;
begin
  if Assigned(APIDLList) and Assigned(FCIDA) then
  begin
    for i := 0 to PIDLCount - 1 do
      APIDLList.Add( PIDLMgr.AppendPIDL(InternalParentPIDL, InternalChildPIDL(i)))
  end;
end;

procedure TCommonShellIDList.AssignPIDLs(APIDLList: TCommonPIDLList);
{ PIDLs[0] must be the Absolute Parent PIDL and the rest single ItemID children }
var
  Count: Integer;
  i: Integer;
  Head: Pointer;
  PIDLLength: Integer;
begin
  Count := 0;
  if Assigned(APIDLList) then
  begin
    { Free previously assigned CIDA }
    if Assigned(FCIDA) then
      FreeMem(FCIDA, CIDASize);
    FCIDA := nil;
    Inc(Count, SizeOf(FCIDA.cidl));
    Inc(Count, SizeOf(FCIDA.aoffset) * (APIDLList.Count));
    for i := 0 to APIDLList.Count - 1 do
      Inc(Count, PIDLMgr.PIDLSize( APIDLList[i]));
    GetMem(FCIDA, Count);
    Head := FCIDA;
    { Head points to the position of the first PIDL }
    Inc(PAnsiChar(Head),  SizeOf(FCIDA.cidl) + (SizeOf(FCIDA.aoffset) * APIDLList.Count));
    { Don't count the absolute parent PIDL }
    FCIDA.cidl := APIDLList.Count - 1;
    for i := 0 to APIDLList.Count - 1 do
    begin
      { Set up the array index to point to the actual PIDL data }
      {$R-}
      FCIDA.aoffset[i] := LongWord(Head-PAnsiChar( CIDA));
      {$R+}
      PIDLLength := PIDLMgr.PIDLSize(APIDLList[i]);
      Move(APIDLList[i]^, Head^, PIDLLength);
      Inc(PAnsiChar(Head), PIDLLength);
    end;
  end;
end;

destructor TCommonShellIDList.Destroy;
begin
  { Free previously assigned CIDA }
  if Assigned(FCIDA) then
    FreeMem(FCIDA, CIDASize);
  inherited;
end;

function TCommonShellIDList.GetCIDASize: integer;
var
  Count: integer;
  i: integer;
begin
  Count := 0;
  if Assigned(FCIDA) then
  begin
    Inc(Count, SizeOf( FCIDA.cidl));
    Inc(Count, SizeOf( FCIDA.aoffset) * (PIDLCount + 1)); // Does't count [0]
    Inc(Count, PIDLMgr.PIDLSize(InternalParentPIDL));
    for i := 0 to PIDLCount - 1 do
      Inc(Count, PIDLMgr.PIDLSize(InternalChildPIDL(i)));
  end;
  Result := Count;
end;

function TCommonShellIDList.GetFormatEtc: TFormatEtc;
begin
  Result.cfFormat := CF_SHELLIDLIST;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL;
end;

function TCommonShellIDList.InternalChildPIDL(Index: integer): PItemIDList;
{ Remember PIDLCount does not count index [0] where the Absolute Parent is     }
begin
  if Assigned(FCIDA) and (Index > -1) and (Index < PIDLCount) then
    Result := PItemIDList( PAnsiChar(FCIDA) + PDWORD(PAnsiChar(@FCIDA^.aoffset)+sizeof(FCIDA^.aoffset[0])*(1+Index))^)
  else
    Result := nil
end;

function TCommonShellIDList.InternalParentPIDL: PItemIDList;
{ Remember PIDLCount does not count index [0] where the Absolute Parent is     }
begin
  if Assigned(FCIDA) then
      Result :=  PItemIDList( PAnsiChar(FCIDA) + FCIDA^.aoffset[0])
  else
    Result := nil
end;

function TCommonShellIDList.LoadFromClipboard: Boolean;
var
  Handle: THandle;
  Ptr: PIDA;
begin
  Result := True;
  Handle := 0;
  if Result then
  begin
    try
      try
        Handle := GetClipboardData(CF_SHELLIDLIST);
        if Handle <> 0 then
        begin
          Ptr := GlobalLock(Handle);
          if Assigned(Ptr) then
          begin
            CIDA := Ptr;
            Result := True;
          end;
        end;
      except
        Result := False;
        raise;
      end;
    finally
      GlobalUnLock(Handle);
    end;
  end
end;

function TCommonShellIDList.LoadFromDataObject(DataObject: IDataObject): Boolean;
var
  Ptr: PIDA;
  StgMedium: TStgMedium;
begin
  Result := False;
  if Assigned(DataObject) then
  begin
    FillChar(StgMedium, SizeOf(StgMedium), #0);
    if Succeeded(DataObject.GetData(GetFormatEtc, StgMedium)) then
    try
      Ptr := GlobalLock(StgMedium.hGlobal);
      try
        if Assigned(Ptr) then
        begin
          CIDA := Ptr;
          Result := True;
        end;
      finally
        GlobalUnLock(StgMedium.hGlobal);
      end;
    finally
      ReleaseStgMedium(StgMedium)
    end;
  end
end;

function TCommonShellIDList.ParentPIDL: PItemIDList;
begin
  Result := PIDLMgr.CopyPIDL( InternalParentPIDL)
end;

function TCommonShellIDList.PIDLCount: integer;
{ indexing is a bit weird.  Index 0 is the Absolute Parent PIDL but it is not }
{ counted in the first byte of the structure.                                 }
begin
  if Assigned(FCIDA) then
    Result := FCIDA^.cidl
  else
    Result := 0
end;

function TCommonShellIDList.RelativePIDL(Index: integer): PItemIDList;
{ Retrieves the single ItemID child by index                                    }
begin
  Result := PIDLMgr.CopyPIDL( InternalChildPIDL(Index))
end;

procedure TCommonShellIDList.RelativePIDLs(APIDLList: TCommonPIDLList);
{ Loads APIDLList with PIDL's stored in the CIDA. ReturnCopy flags if the       }
{ contents will be the origionals or copies created by the PIDLMgr.             }
var
  i: integer;
begin
  if Assigned(APIDLList) and Assigned(FCIDA) then
  begin
    APIDLList.CopyAdd( InternalParentPIDL);
    for i := 0 to PIDLCount - 1 do
      APIDLList.CopyAdd( InternalChildPIDL(i))
  end;
end;

function TCommonShellIDList.SaveToClipboard: Boolean;
var
  DataObject: IDataObject;
begin
  Result := False;
  DataObject := TCommonDataObject.Create;
  if SaveToDataObject(DataObject) then
    Result := Succeeded(OleSetClipboard(DataObject))
end;

function TCommonShellIDList.SaveToDataObject(DataObject: IDataObject): Boolean;
var
  StgMedium: TStgMedium;
  Ptr: PIDA;
begin
  FillChar(StgMedium, SizeOf(StgMedium), #0);

  StgMedium.hGlobal := GlobalAlloc(GPTR, GetCIDASize);
  Ptr := GlobalLock(StgMedium.hGlobal);
  try
    StgMedium.tymed := TYMED_HGLOBAL;
    CopyMemory(Ptr, CIDA, GetCIDASize);
    Result := Succeeded(DataObject.SetData(GetFormatEtc, StgMedium, True))
  finally
    GlobalUnLock(StgMedium.hGlobal);
  end;
end;

procedure TCommonShellIDList.SetCIDA(const Value: PIDA);
var
  TempSize: integer;
begin
  { Free previously assigned CIDA }
  if Assigned(FCIDA) then
  begin
    FreeMem(FCIDA, CIDASize);
    FCIDA := nil;
  end;
  if Value <> nil then
  begin
    { Temporally assign the passed PIDA to the object }
    FCIDA := Value;
    { Get the size of the passed PIDA }
    TempSize := CIDASize;
    { Get memory to make a copy of the passed PIDA }
    GetMem(FCIDA, TempSize);
    { Copy the passed PIDA }
    Move(Value^, FCIDA^, TempSize);
  end;
end;

{ TCommonEnumFormatEtc }

function TCommonEnumFormatEtc.Clone(out Enum: IEnumFormatEtc): HResult;
// Creates a exact copy of the current object.
var
  EnumFormatEtc: TCommonEnumFormatEtc;
begin
  Result := S_OK;   // Think positive
  EnumFormatEtc := TCommonEnumFormatEtc.Create;      // Does not increase COM reference
  if Assigned(EnumFormatEtc) then
  begin
    SetLength(EnumFormatEtc.FFormats, Length(Formats));
    // Make copy of Format info
    Move(FFormats[0], EnumFormatEtc.FFormats[0], Length(Formats) * SizeOf(TFormatEtc));
    EnumFormatEtc.InternalIndex := InternalIndex;
    Enum := EnumFormatEtc as IEnumFormatEtc;   // Sets COM reference to 1
  end else
    Result := E_UNEXPECTED
end;

constructor TCommonEnumFormatEtc.Create;
begin
  inherited Create;
  InternalIndex := 0;
end;

destructor TCommonEnumFormatEtc.Destroy;
begin
  inherited;
end;

function TCommonEnumFormatEtc.Next(celt: Integer; out elt; pceltFetched: PLongint): HResult;
// Another EnumXXXX function.  This function returns the number of objects
// requested by the caller in celt.  The return buffer, elt, is a pointer to an}
// array of, in this case, TFormatEtc structures.  The total number of
// structures returned is placed in pceltFetched.  pceltFetched may be nil if
// celt is only asking for one structure at a time.
var
  i: integer;
begin
  if Assigned(Formats) then
  begin
    i := 0;
    while (i < celt) and (InternalIndex < Length(Formats)) do
    begin
      TeltArray( elt)[i] := Formats[InternalIndex];
      inc(i);
      inc(FInternalIndex);
    end; // while
    if assigned(pceltFetched) then
      pceltFetched^ := i;
    if i = celt then
      Result := S_OK
    else
      Result := S_FALSE
  end else
    Result := E_UNEXPECTED
end;

function TCommonEnumFormatEtc.Reset: HResult;
begin
  InternalIndex := 0;
  Result := S_OK
end;

function TCommonEnumFormatEtc.Skip(celt: Integer): HResult;
// Allows the caller to skip over unwanted TFormatEtc structures.  Simply adds
// celt to the index as long as it does not skip past the last structure in
// the list.
begin
  if Assigned(Formats) then
  begin
    if InternalIndex + celt < Length(Formats) then
    begin
      InternalIndex := InternalIndex + celt;
      Result := S_OK
    end else
      Result := S_FALSE
  end else
    Result := E_UNEXPECTED
end;

procedure TCommonEnumFormatEtc.SetFormatLength(Size: Integer);
begin
  SetLength(FFormats, Size)
end;

{ TCommonDataObject }
class function TCommonDataObject.NewInstance: TObject;
begin
  Result := inherited NewInstance;
  TCommonDataObject(Result).FRefCount := 1;
end;


function TCommonDataObject.AssignDragImage(Image: TBitmap;
  HotSpot: TPoint; TransparentColor: TColor): Boolean;
//
// Stores the Bitmap in the IDataObject to support the IDragSourceHelper drag image
// in Win2K and above.
//
var
  DragSourceHelper: IDragSourceHelper;
  SHDragImage: TSHDragImage;
begin
  Result := False;
  // NT can't swallow this CoCreateInstance call
  if not IsWinNT4 then
  begin
    if Succeeded(CoCreateInstance(CLSID_DragDropHelper, nil, CLSCTX_INPROC_SERVER, IID_IDragSourceHelper, DragSourceHelper)) then
    begin
      FillChar(SHDragImage, SizeOf(SHDragImage), #0);

      SHDragImage.sizeDragImage.cx := Image.Width;
      SHDragImage.sizeDragImage.cy := Image.Height;
      SHDragImage.ptOffset := HotSpot;
      SHDragImage.ColorRef := ColorToRGB(TransparentColor);
      SHDragImage.hbmpDragImage := CopyImage(Image.Handle, IMAGE_BITMAP, Image.Width,
        Image.Height, LR_COPYRETURNORG);
      if SHDragImage.hbmpDragImage <> 0 then
        if Succeeded(DragSourceHelper.InitializeFromBitmap(SHDragImage, Self as IDataObject)) then
          Result := True
        else
          DeleteObject(SHDragImage.hbmpDragImage);
    end
  end
end;

function TCommonDataObject.CanonicalIUnknown(TestUnknown: IUnknown): IUnknown;
// Uses COM object identity: An explicit call to the IUnknown::QueryInterface
// method, requesting the IUnknown interface, will always return the same
// pointer.
begin
  if Assigned(TestUnknown) then
  begin
    if CommonSupports(TestUnknown, IUnknown, Result) then
      IUnknown(Result)._Release // Don't actually need it just need the pointer value
    else
      Result := TestUnknown
  end else
    Result := TestUnknown
end;

constructor TCommonDataObject.Create(IsReferenceCounted: Boolean = True);
begin
  inherited Create;
  FReferenceCounted := IsReferenceCounted;
end;

function TCommonDataObject.DAdvise(const formatetc: TFormatEtc; advf: Integer;
  const advSink: IAdviseSink; out dwConnection: Integer): HResult;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.DAdvise');
  {$ENDIF}
  if not Assigned(AdviseHolder) then
    CreateDataAdviseHolder(FAdviseHolder);
  if Assigned(FAdviseHolder) then
    Result := AdviseHolder.Advise(Self as IDataObject, formatetc, advf, advSink, dwConnection)
  else
    Result := OLE_E_ADVISENOTSUPPORTED;
end;

destructor TCommonDataObject.Destroy;
begin
  inherited;
end;

function TCommonDataObject.GetObj: TObject;
begin
  Result := Self
end;

function TCommonDataObject.MultiFolder: Boolean;
begin
  Result := IsMultiFolder
end;

function TCommonDataObject.MultiFolderVerified: Boolean;
begin
  Result := IsMultiFolderVerified
end;

function TCommonDataObject.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.QueryInterface: ' + GUIDToInterfaceStr(IID));
  {$ENDIF}
  if GetInterface(IID, Obj) then
    Result := 0
  else
    Result := E_NOINTERFACE;
end;

function TCommonDataObject._AddRef: Integer;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject._AddRef');
  {$ENDIF}
  if ReferenceCounted then
    Result := InterlockedIncrement(FRefCount)
  else
    Result := -1
end;

function TCommonDataObject._Release: Integer;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject._Release');
  {$ENDIF}
  if ReferenceCounted then
  begin
    Result := InterlockedDecrement(FRefCount);
    if Result = 0 then
      Destroy;
  end else
    Result := -1
end;

procedure TCommonDataObject.AfterConstruction;
begin
// Release the constructor's implicit refcount
  InterlockedDecrement(FRefCount);
end;

procedure TCommonDataObject.BeforeDestruction;
begin
//  if RefCount <> 0 then
//    Error(reInvalidPtr);
end;

procedure TCommonDataObject.DoGetCustomFormats(dwDirection: Integer; var Formats: TFormatEtcArray);
begin
  
end;

procedure TCommonDataObject.DoOnGetData(const FormatEtcIn: TFormatEtc;
  var Medium: TStgMedium; var Handled: Boolean);
begin
  if Assigned(FOnGetData) then
    OnGetData(Self, FormatEtcIn, Medium, Handled);
end;

procedure TCommonDataObject.DoOnQueryGetData(
  const FormatEtcIn: TFormatEtc; var FormatAvailable: Boolean; var Handled: Boolean);
begin
  if Assigned(FOnQueryGetData) then
    OnQueryGetData(Self, FormatEtcIn, FormatAvailable, Handled);
end;

function TCommonDataObject.DUnadvise(dwConnection: Integer): HResult;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.DUnadvise');
  {$ENDIF}
  if Assigned(AdviseHolder) then
    Result := AdviseHolder.Unadvise(dwConnection)
  else
    Result := OLE_E_ADVISENOTSUPPORTED;
end;

function TCommonDataObject.EnumDAdvise(out enumAdvise: IEnumStatData): HResult;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.EnumDAdvise');
  {$ENDIF}
  if Assigned(AdviseHolder) then
    Result := AdviseHolder.EnumAdvise(enumAdvise)
  else
    Result := OLE_E_ADVISENOTSUPPORTED;
end;

function TCommonDataObject.EnumFormatEtc(dwDirection: Integer;
  out enumFormatEtc: IEnumFormatEtc): HResult;
// Called when DoDragDrop wants to know what clipboard formats are supported
// by Enumerating the TFormatEtc array through an IEnumFormatEtc object.
var
  LocalEnumFormatEtc: TCommonEnumFormatEtc;
  i: integer;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.EnumFormatEtc');
  {$ENDIF}
  // Always return an object even if it is empty
  LocalEnumFormatEtc := TCommonEnumFormatEtc.Create;
  // Get the reference count in sync
  enumFormatEtc := LocalEnumFormatEtc as IEnumFormatEtc;
{  if Assigned(Formats) then
  begin        }
    Result := S_OK;
    if dwDirection = DATADIR_GET then
    begin
      // Copy the internal supported Formats for the EnumFormatEtc
      SetLength(LocalEnumFormatEtc.FFormats, Length(Formats));
      for i := 0 to Length(Formats) - 1 do
        LocalEnumFormatEtc.Formats[i] := Formats[i].FormatEtc;

      // Now copy any custom formats
      DoGetCustomFormats(dwDirection, LocalEnumFormatEtc.FFormats);

      if not Assigned(enumFormatEtc) then
        Result := E_OUTOFMEMORY
    end else
    begin
      enumFormatEtc := nil;
      Result := E_NOTIMPL;
    end;
 { end else
  begin
    if dwDirection = DATADIR_GET then
      Result := S_OK
    else
      Result := E_NOTIMPL
  end }
end;

function TCommonDataObject.EqualFormatEtc(FormatEtc1, FormatEtc2: TFormatEtc): Boolean;
begin
  Result := (FormatEtc1.cfFormat = FormatEtc2.cfFormat) and
            (FormatEtc1.ptd = FormatEtc2.ptd) and
            (FormatEtc1.dwAspect = FormatEtc2.dwAspect) and
            (FormatEtc1.lindex = FormatEtc2.lindex) and
            (FormatEtc1.tymed = FormatEtc2.tymed)
end;

function TCommonDataObject.FindFormatEtc(TestFormatEtc: TFormatEtc): integer;
var
  i: integer;
  Found: Boolean;
begin
  i := 0;
  Found := False;
  Result := -1;
  while (i < Length(FFormats)) and not Found do
  begin
    Found := EqualFormatEtc(Formats[i].FormatEtc, TestFormatEtc);
    if Found then
      Result := i;
    Inc(i);
  end
end;

function TCommonDataObject.GetCanonicalFormatEtc(const formatetc: TFormatEtc;
  out formatetcOut: TFormatEtc): HResult;
// Since we do not have two TFormatEtcs that return the same type of data we can
// ingore this function.  It is only for TFormatEtc structures that will return
// the exact same data if each is called.  This could happen if the data is
// target dependant and the target can handle both types of data.  This keeps
// the target from asking for redundant information.
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.GetCanonicalFormatEtc');
  {$ENDIF}
  formatetcOut.ptd := nil;
  Result := E_NOTIMPL;
end;

function TCommonDataObject.GetData(const FormatEtcIn: TFormatEtc;
  out Medium: TStgMedium): HResult;
// This is the workhorse of the functions.  It looks at the clipboard format
// the IDropTarget wants, makes sure we can support it.  If supported then see
// if it is owned by the object or the program will supply the data.
var
  Handled: Boolean;
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  ClipName: array[0..128] of AnsiChar;
  {$ENDIF}
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  if GetClipboardFormatNameA(FormatEtcIn.cfFormat, ClipName, 128) > 0 then
    SendDebug('TCommonDataObject.GetData: ' + ClipName)
  else
    SendDebug('TCommonDataObject.GetData: ' + IntToStr(FormatEtcIn.cfFormat));
  {$ENDIF}
  Result := E_UNEXPECTED;
  Handled := False;
  DoOnGetData(FormatEtcIn, Medium, Handled);
  if not Handled then
  begin
  if Assigned(Formats) then
    begin
      { Do we support this type of Data? }
      Result := QueryGetData(FormatEtcIn);
      if Result = S_OK then
      begin
        // If the data is owned by the IDataObject just retrieve and return it.
        if RetrieveOwnedStgMedium(FormatEtcIn, Medium) = E_INVALIDARG then
        { This data is defined by the Object Inspector or a custom format need to }
        { Retrive it from the DragSourceManager                                   }
          if not GetUserData(FormatEtcIn, Medium) then
            Result := E_UNEXPECTED
      end
    end
  end else
    Result := S_OK
end;

function TCommonDataObject.GetDataHere(const formatetc: TFormatEtc;
  out medium: TStgMedium): HResult;
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  SendDebug('TCommonDataObject.GetDataHere');
  {$ENDIF}
  Result := E_NOTIMPL;
end;

function TCommonDataObject.GetUserData(Format: TFormatEtc; var StgMedium: TStgMedium): Boolean;
begin
  Result := False;
end;

function TCommonDataObject.HGlobalClone(HGlobal: THandle): THandle;
// Returns a global memory block that is a copy of the passed memory block.
var
  Size: LongWord;
  Data, NewData: PAnsiChar;
begin
  Size := GlobalSize(HGlobal);
  Result := GlobalAlloc(GPTR, Size);
  Data := GlobalLock(hGlobal);
  try
    NewData := GlobalLock(Result);
    try
      Move(Data, NewData, Size);
    finally
      GlobalUnLock(Result);
    end
  finally
    GlobalUnLock(hGlobal)
  end
end;

function TCommonDataObject.LoadGlobalBlock(Format: TClipFormat;
  var MemoryBlock: Pointer): Boolean;
var
  FormatEtc: TFormatEtc;
  StgMedium: TStgMedium;
  GlobalObject: Pointer;
begin
  Result := False;

  FormatEtc.cfFormat := Format;
  FormatEtc.ptd := nil;
  FormatEtc.dwAspect := DVASPECT_CONTENT;
  FormatEtc.lindex := -1;
  FormatEtc.tymed := TYMED_HGLOBAL;

  if Succeeded(QueryGetData(FormatEtc)) and Succeeded(GetData(FormatEtc, StgMedium)) then
  begin
    MemoryBlock := AllocMem( GlobalSize(StgMedium.hGlobal));
    GlobalObject := GlobalLock(StgMedium.hGlobal);
    try
      if Assigned(MemoryBlock) and Assigned(GlobalObject) then
      begin
        Move(GlobalObject^, MemoryBlock^, GlobalSize(StgMedium.hGlobal));
      end
    finally
      GlobalUnLock(StgMedium.hGlobal);
    end
  end;
end;

function TCommonDataObject.QueryGetData(const formatetc: TFormatEtc): HResult;
// This function allows the IDragTarget to see if we can possibly support some
// type of data transfer.
var
  i: integer;
  FormatAvailable, Handled: Boolean;
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  ClipName: array[0..128] of AnsiChar;
  {$ENDIF}
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  if GetClipboardFormatNameA(formatetc.cfFormat, ClipName, 128) > 0 then
    SendDebug('TCommonDataObject.QueryGetData: ' + ClipName)
  else
    SendDebug('TCommonDataObject.QueryGetData: ' + IntToStr(formatetc.cfFormat));
  {$ENDIF}
  Handled := False;
  FormatAvailable := False;
  DoOnQueryGetData(FormatEtc, FormatAvailable, Handled);
  if Handled then
  begin
    if FormatAvailable then
      Result := S_OK
    else
      Result := DV_E_FORMATETC
  end else
  begin
    if not FormatAvailable then
    begin
      if Assigned(Formats) then
      begin
        i := 0;
        Result := DV_E_FORMATETC;
        while (i < Length(Formats)) and (Result = DV_E_FORMATETC)  do
        begin
          if Formats[i].FormatEtc.cfFormat = formatetc.cfFormat then
          begin
            if (Formats[i].FormatEtc.dwAspect = formatetc.dwAspect) then
            begin
              if (Formats[i].FormatEtc.tymed and formatetc.tymed <> 0) then
                Result := S_OK
              else
                Result := DV_E_TYMED;
            end else
              Result := DV_E_DVASPECT;
          end else
            Result := DV_E_FORMATETC;
          Inc(i)
        end
      end else
        Result := E_UNEXPECTED;
    end else
      Result := S_OK
  end
end;

function TCommonDataObject.RetrieveOwnedStgMedium(Format: TFormatEtc; var StgMedium: TStgMedium): HRESULT;
var
  i: integer;
begin
  Result := E_INVALIDARG;
  i := FindFormatEtc(Format);
  if (i > -1) and Formats[i].OwnedByDataObject then
    Result := StgMediumIncRef(Formats[i].StgMedium, StgMedium, False)
end;

function TCommonDataObject.SaveGlobalBlock(Format: TClipFormat;
  MemoryBlock: Pointer; MemoryBlockSize: integer): Boolean;
var
  FormatEtc: TFormatEtc;
  StgMedium: TStgMedium;
  GlobalObject: Pointer;
begin
  FormatEtc.cfFormat := Format;
  FormatEtc.ptd := nil;
  FormatEtc.dwAspect := DVASPECT_CONTENT;
  FormatEtc.lindex := -1;
  FormatEtc.tymed := TYMED_HGLOBAL;

  StgMedium.tymed := TYMED_HGLOBAL;
  StgMedium.unkForRelease := nil;
  StgMedium.hGlobal := GlobalAlloc(GHND or GMEM_SHARE, MemoryBlockSize);
  GlobalObject := GlobalLock(StgMedium.hGlobal);
  try
    try
      Move(MemoryBlock^, GlobalObject^, MemoryBlockSize);
      Result := Succeeded( SetData(FormatEtc, StgMedium, True))
    except
      Result := False;
    end
  finally
    GlobalUnLock(StgMedium.hGlobal);
  end
end;

function TCommonDataObject.SetData(const formatetc: TFormatEtc; var medium: TStgMedium; fRelease: BOOL): HResult;
// Allows dynamic adding to the IDataObject during its existance.  Most noteably
// it is used to implement IDropSourceHelper in win2k
var
  Index: integer;
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  ClipName: array[0..128] of AnsiChar;
  {$ENDIF}
begin
  {$IFDEF GX_DEBUG_COMMONDATAOBJECT}
  if GetClipboardFormatNameA(formatetc.cfFormat, ClipName, 128) > 0 then
    SendDebug('TCommonDataObject.SetData: ' + ClipName)
  else
    SendDebug('TCommonDataObject.SetData: ' + IntToStr(formatetc.cfFormat));
  {$ENDIF}
  // See if we already have a format of that type available.
  Index := FindFormatEtc(FormatEtc);
  if Index > - 1 then
  begin
    // Yes we already have that format type stored.  Just use the TCommonClipboardFormat
    // in the List after releasing the data
    ReleaseStgMedium(Formats[Index].StgMedium);
    FillChar(Formats[Index].StgMedium, SizeOf(Formats[Index].StgMedium), #0);
  end else
  begin
    // It is a new format so create a new TDataObjectInfo record and store it in
    // the Format array
    SetLength(FFormats, Length(Formats) + 1);
    Formats[Length(Formats) - 1].FormatEtc := FormatEtc;
    Index := Length(Formats) - 1;
  end;
  // The data is owned by the TCommonClipboardFormat object
  Formats[Index].OwnedByDataObject := True;

  if fRelease then
  begin
    // We are simply being given the data and we take control of it.
    Formats[Index].StgMedium := Medium;
    Result := S_OK
  end else
    // We need to reference count or copy the data and keep our own references
    // to it.
    Result := StgMediumIncRef(Medium, Formats[Index].StgMedium, True);

    // Can get a circular reference if the client calls GetData then calls
    // SetData with the same StgMedium.  Because the unkForRelease and for
    // the IDataObject can be marshalled it is necessary to get pointers that
    // can be correctly compared.
    // See the IDragSourceHelper article by Raymond Chen at MSDN.
    if Assigned(Formats[Index].StgMedium.unkForRelease) then
    begin
      if CanonicalIUnknown(Self) =
        CanonicalIUnknown(IUnknown( Formats[Index].StgMedium.unkForRelease)) then
      begin
        IUnknown( Formats[Index].StgMedium.unkForRelease)._Release;
        Formats[Index].StgMedium.unkForRelease := nil
      end;
    end;
  // Tell all registered advice sinks about the data change.
  if Assigned(AdviseHolder) then
    AdviseHolder.SendOnDataChange(Self as IDataObject, 0, 0);
end;

function TCommonDataObject.StgMediumIncRef(const InStgMedium: TStgMedium;
  var OutStgMedium: TStgMedium; CopyInMedium: Boolean): HRESULT;
// This function increases the reference count of the passed StorageMedium in a
// variety of ways depending on the value of CopyInMedium.
// InStgMedium is the data that is requested a copy of, OutStgMedium is the data that
// we are to return either a copy of or increase the IDataObject's reference and
// send ourselves back as the data (unkForRelease). The InStgMedium is usually
// the result of a call to find a particular FormatEtc that has been stored
// locally through a call to SetData.     If CopyInMedium is not true we
// already have a local copy of the data when the SetData function was called
// (during that call the CopyInMedium must be true).  Then as the caller asks
// for the data through GetData we do not have to make copy of the data for the
// caller only to have them destroy it then need us to copy it again if
// necessary.  This way we increase the reference count to ourselves and pass
// the STGMEDIUM structure initially stored in SetData.  This way when the
// caller frees the structure it sees the unkForRelease is not nil and calls
// Release on the object instead of destroying the actual data.
begin
  Result := S_OK;
  // Simply copy all fields to start with.
  OutStgMedium := InStgMedium;
  case InStgMedium.tymed of
    TYMED_HGLOBAL:
      begin
        if CopyInMedium then
        begin
          // Generate a unique copy of the data passed
          OutStgMedium.hGlobal := HGlobalClone(InStgMedium.hGlobal);
          if OutStgMedium.hGlobal = 0 then
            Result := E_OUTOFMEMORY
        end else
          // Don't generate a copy just use ourselves and the copy previoiusly saved
          OutStgMedium.unkForRelease := Pointer(Self as IDataObject) // Does increase RefCount
      end;
    TYMED_FILE:
      begin
        if CopyInMedium then
        begin
          OutStgMedium.lpszFileName := CoTaskMemAlloc(lstrLenW(InStgMedium.lpszFileName));
          MoveMemory(PWideChar(OutStgMedium.lpszFileName), PWideChar(InStgMedium.lpszFileName), lstrlenW(InStgMedium.lpszFileName) * 2);
        end else
          OutStgMedium.unkForRelease := Pointer(Self as IDataObject) // Does increase RefCount
      end;
    TYMED_ISTREAM:
      // Simply increase the reference so the stream object
      // Note here stm is a pointer to the auto reference counting won't work and
      // we have to call _AddRef explicitly
      IUnknown( OutStgMedium.stm)._AddRef;
    TYMED_ISTORAGE:
      // Simply increase the reference so the storage object
      // Note here stm is a pointer to the auto reference counting won't work and
      // we have to call _AddRef explicitly
      IUnknown( OutStgMedium.stg)._AddRef;
    TYMED_GDI:
      if not CopyInMedium then
      // Don't generate a copy just use ourselves and the copy previoiusly saved data
        OutStgMedium.unkForRelease := Pointer(Self as IDataObject) // Does not increase RefCount
     else
       Result := DV_E_TYMED; // Don't know how to copy GDI objects right now
    TYMED_MFPICT:
      if not CopyInMedium then
        OutStgMedium.unkForRelease := Pointer(Self as IDataObject) // Does not increase RefCount
      else
        Result := DV_E_TYMED; // Don't know how to copy MetaFile objects right now
    TYMED_ENHMF:
      if not CopyInMedium then
        { Don't generate a copy just use ourselves and the copy previoiusly saved data }
        OutStgMedium.unkForRelease := Pointer(Self as IDataObject) // Does not increase RefCount
      else
        Result := DV_E_TYMED; // Don't know how to copy enhanced metafiles objects right now
  else
    Result := DV_E_TYMED
  end;

  // I still have to do this. The Compiler will call _Release on the above Self as IDataObject
  // casts which is not what is necessary.  The DataObject is released correctly.
  if Assigned(OutStgMedium.unkForRelease) and (Result = S_OK) then
    IUnknown(OutStgMedium.unkForRelease)._AddRef
end;

procedure TCommonDataObject.SetMultiFolder(IsSet: Boolean);
begin
  IsMultiFolder := IsSet
end;

procedure TCommonDataObject.SetMultiFolderVerified(IsVerified: Boolean);
begin
  IsMultiFolderVerified := IsVerified
end;

{ TFileGroupDescriptorA }
destructor TFileGroupDescriptorA.Destroy;
begin
  FreeAndNil(FStream);
  inherited Destroy;
end;


procedure TFileGroupDescriptorA.AddFileDescriptor(
  FileDescriptor: TFileDescriptorA);
begin
  SetLength(FFileDescriptors, Length(FFileDescriptors) + 1);
  FFileDescriptors[Length(FFileDescriptors) - 1] := FileDescriptor;
end;

procedure TFileGroupDescriptorA.DeleteFileDescriptor(Index: integer);
var
  i: Integer;
begin
  for i := Index to Length(FFileDescriptors) - 1 do
    FileDescriptor[i] := FileDescriptor[i+1];
  SetLength(FFileDescriptors, Length(FFileDescriptors) - 1);
end;

function TFileGroupDescriptorA.FillDescriptor(FileName: AnsiString): TFileDescriptorA;
begin
  FillChar(Result, SizeOf(Result), #0);
  StrCopy(Result.cFileName, PAnsiChar(FileName));
end;

function TFileGroupDescriptorA.GetDescriptorCount: Integer;
begin
  Result := Length(FFileDescriptors)
end;

function TFileGroupDescriptorA.GetFileDescriptorA(Index: Integer): TFileDescriptorA;
begin
  FillChar(Result, SizeOf(Result), #0);
  if (Index > -1) and (Index < Length(FFileDescriptors)) then
    Result := FFileDescriptors[Index]
end;

function TFileGroupDescriptorA.GetFileStream(const DataObject: IDataObject;
  FileIndex: Integer): TStream;
var
 Format: TFormatEtc;
 Medium: TStgMedium;
 PMem : Pointer;
 Ok : boolean;
begin
  if Assigned(Stream) then
    FreeAndNil(FStream);

  if Assigned(DataObject) and (FileIndex > -1) and (FileIndex < DescriptorCount) then
  begin
    Format.cfFormat := CF_FILECONTENTS;
    Format.ptd := nil;
    Format.dwAspect := DVASPECT_CONTENT;
    Format.lindex := FileIndex;
    Format.tymed := TYMED_ISTREAM;
    Ok := Succeeded(DataObject.GetData(Format, Medium));
    if Ok then
    begin
      if Medium.tymed = TYMED_ISTREAM then
      begin
        FStream := TOLEStream.Create(IStream(Medium.stm));
        ReleaseStgMedium(Medium);
       end else
         Ok := false;
    end;
    if not Ok then
    begin
      Format.tymed := TYMED_HGLOBAL;
      if Succeeded(DataObject.GetData(Format, Medium)) then
      begin
        FStream := TMemoryStream.Create;
        PMem := GlobalLock(Medium.hGlobal);
        FStream.Size := GlobalSize(Medium.hglobal);
        MoveMemory(TMemoryStream(FStream).Memory, PMem, FStream.Size);
        GlobalUnLock(Medium.hGlobal);
        ReleaseStgMedium(Medium);
      end
    end
  end;
  Result := Stream;
end;

function TFileGroupDescriptorA.GetFormatEtc: TFormatEtc;
begin
  Result := FileDescriptorAFormat 
end;

procedure TFileGroupDescriptorA.LoadFileGroupDestriptor(FileGroupDiscriptor: PFileGroupDescriptorA);
var
  i: Cardinal;
begin
  if Assigned(FileGroupDiscriptor) then
  begin
    SetLength(FFileDescriptors, FileGroupDiscriptor.cItems);
    for i := 0 to FileGroupDiscriptor.cItems - 1 do
    begin
      FFileDescriptors[i] := FileGroupDiscriptor.fgd[i]
    end
  end else
    FFileDescriptors := nil;
end;

function TFileGroupDescriptorA.LoadFromClipboard: Boolean;
var
  DataObject: IDataObject;
begin
  Result := False;
  if Succeeded(OleGetClipboard(DataObject)) then
    Result := LoadFromDataObject(DataObject);
end;

function TFileGroupDescriptorA.LoadFromDataObject(DataObject: IDataObject): Boolean;
var
  GroupDescriptor: PFileGroupDescriptorA;
  Medium: TStgMedium;
  i: Integer;
begin
  Result := False;
  if Succeeded(DataObject.GetData(GetFormatEtc, Medium)) then
  begin
    GroupDescriptor := GlobalLock(Medium.hGlobal);
    try
      for i := 0 to GroupDescriptor^.cItems - 1 do
        AddFileDescriptor(GroupDescriptor^.fgd[i])
    finally
      GlobalUnlock(Medium.hGlobal);
      ReleaseStgMedium(Medium);
      Result := True
    end
  end
end;

function TFileGroupDescriptorA.SaveToClipboard: Boolean;
var
  DataObject: IDataObject;
begin
  Result := False;
  DataObject := TCommonDataObject.Create;
  if SaveToDataObject(DataObject) then
    Result := Succeeded(OleSetClipboard(DataObject))
end;

function TFileGroupDescriptorA.SaveToDataObject(DataObject: IDataObject): Boolean;
var
  Mem: THandle;
  GroupDescriptor: PFileGroupDescriptorA;
  Medium: TStgMedium;
  Format: TFormatEtc;
begin
  Result := False;
  if Assigned(DataObject) and (DescriptorCount > 0) then
  begin
    Mem := GlobalAlloc(GHND, DescriptorCount * SizeOf(TFileDescriptorA) + SizeOf(GroupDescriptor.cItems));
    GroupDescriptor := GlobalLock(Mem);
    try
      GroupDescriptor.cItems := DescriptorCount;
      CopyMemory(@GroupDescriptor^.fgd[0], @FFileDescriptors[0], DescriptorCount * SizeOf(TFileDescriptorA));
    finally
      GlobalUnlock(Mem)
    end;
    FillChar(Medium, SizeOf(Medium), #0);
    Medium.tymed := TYMED_HGLOBAL;
    Medium.hGlobal := Mem;

    DataObject.SetData(GetFormatEtc, Medium, True);

    Medium.tymed := TYMED_ISTREAM;
    Medium.stm := nil;

    Format.cfFormat := CF_FILECONTENTS;
    Format.ptd := nil;
    Format.dwAspect := DVASPECT_CONTENT;
    Format.lindex := -1;
    Format.tymed := TYMED_ISTREAM;
    DataObject.SetData(Format, Medium, True);
  end
end;

procedure TFileGroupDescriptorA.SetFileDescriptor(Index: Integer; const Value: TFileDescriptorA);
begin
  if (Index > -1) and (Index < Length(FFileDescriptors)) then
    FFileDescriptors[Index] := Value
end;

{ TFileGroupDescriptorW }
destructor TFileGroupDescriptorW.Destroy;
begin
  FreeAndNil(FStream);
  inherited Destroy;
end;


procedure TFileGroupDescriptorW.AddFileDescriptor(
  FileDescriptor: TFileDescriptorW);
begin
  SetLength(FFileDescriptors, Length(FFileDescriptors) + 1);
  FFileDescriptors[Length(FFileDescriptors) - 1] := FileDescriptor;
end;

procedure TFileGroupDescriptorW.DeleteFileDescriptor(Index: integer);
var
  i: Integer;
begin
  for i := Index to Length(FFileDescriptors) - 1 do
    FileDescriptor[i] := FileDescriptor[i+1];
  SetLength(FFileDescriptors, Length(FFileDescriptors) - 1);
end;

function TFileGroupDescriptorW.FillDescriptor(FileName: WideString): TFileDescriptorW;
begin
  FillChar(Result, SizeOf(Result), #0);
  StrCopyW(Result.cFileName, PWideChar(FileName));
end;

function TFileGroupDescriptorW.GetDescriptorCount: Integer;
begin
  Result := Length(FFileDescriptors)
end;

function TFileGroupDescriptorW.GetFileDescriptorW(Index: Integer): TFileDescriptorW;
begin
  FillChar(Result, SizeOf(Result), #0);
  if (Index > -1) and (Index < Length(FFileDescriptors)) then
    Result := FFileDescriptors[Index]
end;

function TFileGroupDescriptorW.GetFileStream(const DataObject: IDataObject; FileIndex: Integer): TStream;
var
 Format: TFormatEtc;
 Medium: TStgMedium;
 PMem : Pointer;
 Ok : boolean;
begin
  if Assigned(Stream) then
    FreeAndNil(FStream);

  if Assigned(DataObject) and (FileIndex > -1) and (FileIndex < DescriptorCount) then
  begin
    Format.cfFormat := CF_FILECONTENTS;
    Format.ptd := nil;
    Format.dwAspect := DVASPECT_CONTENT;
    Format.lindex := FileIndex;
    Format.tymed := TYMED_ISTREAM;
    Ok := Succeeded(DataObject.GetData(Format, Medium));
    if Ok then
    begin
      if Medium.tymed = TYMED_ISTREAM then
      begin
        FStream := TOLEStream.Create(IStream(Medium.stm));
        ReleaseStgMedium(Medium);
       end else
         Ok := false;
    end;
    if not Ok then
    begin
      Format.tymed := TYMED_HGLOBAL;
      if Succeeded(DataObject.GetData(Format, Medium)) then
      begin
        FStream := TMemoryStream.Create;
        PMem := GlobalLock(Medium.hGlobal);
        FStream.Size := GlobalSize(Medium.hglobal);
        MoveMemory(TMemoryStream(FStream).Memory, PMem, FStream.Size);
        GlobalUnLock(Medium.hGlobal);
        ReleaseStgMedium(Medium);
      end
    end
  end;
  Result := Stream;
end;

function TFileGroupDescriptorW.GetFormatEtc: TFormatEtc;
begin
  Result := FileDescriptorWFormat 
end;

procedure TFileGroupDescriptorW.LoadFileGroupDestriptor(FileGroupDiscriptor: PFileGroupDescriptorW);
var
  i: Cardinal;
begin
  if Assigned(FileGroupDiscriptor) then
  begin
    SetLength(FFileDescriptors, FileGroupDiscriptor.cItems);
    for i := 0 to FileGroupDiscriptor.cItems - 1 do
    begin
      FFileDescriptors[i] := FileGroupDiscriptor.fgd[i]
    end
  end else
    FFileDescriptors := nil;
end;

function TFileGroupDescriptorW.LoadFromClipboard: Boolean;
var
  DataObject: IDataObject;
begin
  Result := False;
  if Succeeded(OleGetClipboard(DataObject)) then
    Result := LoadFromDataObject(DataObject);
end;

function TFileGroupDescriptorW.LoadFromDataObject(DataObject: IDataObject): Boolean;
var
  GroupDescriptor: PFileGroupDescriptorW;
  Medium: TStgMedium;
  i: Integer;
begin
  Result := False;
  if Succeeded(DataObject.GetData(GetFormatEtc, Medium)) then
  begin
    GroupDescriptor := GlobalLock(Medium.hGlobal);
    try
      for i := 0 to GroupDescriptor^.cItems - 1 do
        AddFileDescriptor(GroupDescriptor^.fgd[i])
    finally
      GlobalUnlock(Medium.hGlobal);
      ReleaseStgMedium(Medium);
      Result := True;
    end
  end
end;

function TFileGroupDescriptorW.SaveToClipboard: Boolean;
var
  DataObject: IDataObject;
begin
  Result := False;
  DataObject := TCommonDataObject.Create;
  if SaveToDataObject(DataObject) then
    Result := Succeeded(OleSetClipboard(DataObject))
end;

function TFileGroupDescriptorW.SaveToDataObject(DataObject: IDataObject): Boolean;
var
  Mem: THandle;
  GroupDescriptor: PFileGroupDescriptorW;
  Medium: TStgMedium;
  Format: TFormatEtc;
begin
  Result := False;
  if Assigned(DataObject) and (DescriptorCount > 0) then
  begin
    Mem := GlobalAlloc(GHND, DescriptorCount * SizeOf(TFileDescriptorW) + SizeOf(GroupDescriptor.cItems));
    GroupDescriptor := GlobalLock(Mem);
    try
      GroupDescriptor.cItems := DescriptorCount;
      CopyMemory(@GroupDescriptor^.fgd[0], @FFileDescriptors[0], DescriptorCount * SizeOf(TFileDescriptorW));
    finally
      GlobalUnlock(Mem)
    end;
    FillChar(Medium, SizeOf(Medium), #0);
    Medium.tymed := TYMED_HGLOBAL;
    Medium.hGlobal := Mem;

    DataObject.SetData(GetFormatEtc, Medium, True);

    Medium.tymed := TYMED_ISTREAM;
    Medium.stm := nil;

    Format.cfFormat := CF_FILECONTENTS;
    Format.ptd := nil;
    Format.dwAspect := DVASPECT_CONTENT;
    Format.lindex := -1;
    Format.tymed := TYMED_ISTREAM;
    DataObject.SetData(Format, Medium, True);
  end
end;

procedure TFileGroupDescriptorW.SetFileDescriptor(Index: Integer; const Value: TFileDescriptorW);
begin
  if (Index > -1) and (Index < Length(FFileDescriptors)) then
    FFileDescriptors[Index] := Value
end;

{ TCommonInShellDragLoop }
function TCommonInShellDragLoop.GetFormatEtc: TFormatEtc;
begin
  Result.cfFormat := CF_INDRAGLOOP;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := TYMED_HGLOBAL
end;

function TCommonInShellDragLoop.LoadFromDataObject(DataObject: IDataObject): Boolean;
var
  StgMedium: TStgMedium;
  DataPtr: PCardinal;
begin
  Result := False;
  FillChar(StgMedium, SizeOf(StgMedium), #0);

  if Succeeded(DataObject.GetData(GetFormatEtc, StgMedium)) then
  try
    DataPtr := GlobalLock(StgMedium.hGlobal);
    try
      if Assigned(DataPtr) then
      begin
        Data := DataPtr^;
        Result := True;
      end;
    finally
      GlobalUnLock(StgMedium.hGlobal);
    end
  finally
    ReleaseStgMedium(StgMedium)
  end
end;

function TCommonInShellDragLoop.SaveToDataObject(DataObject: IDataObject): Boolean;
var
  StgMedium: TStgMedium;
  Ptr: PCardinal;
begin
  FillChar(StgMedium, SizeOf(StgMedium), #0);
  StgMedium.hGlobal := GlobalAlloc(GPTR, SIZE_SHELLDRAGLOOPDATA);
  Ptr := GlobalLock(StgMedium.hGlobal);
  try
    StgMedium.tymed := TYMED_HGLOBAL;
    Ptr^ := FData;
    Result := Succeeded(DataObject.SetData(GetFormatEtc, StgMedium, True))
  finally
    GlobalUnLock(StgMedium.hGlobal);
  end
end;

initialization
  CF_SHELLIDLIST := RegisterClipboardFormat(CFSTR_SHELLIDLIST);
  CF_LOGICALPERFORMEDDROPEFFECT := RegisterClipboardFormat(CFSTR_LOGICALPERFORMEDDROPEFFECT);
  CF_PREFERREDDROPEFFECT := RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
  CF_PERFORMEDDROPEFFECT := RegisterClipboardFormat(CFSTR_PERFORMEDDROPEFFECT);
  CF_PASTESUCCEEDED := RegisterClipboardFormat(CFSTR_PASTESUCCEEDED);
  CF_INDRAGLOOP := RegisterClipboardFormat(CFSTR_INDRAGLOOP);
  CF_SHELLIDLISTOFFSET := RegisterClipboardFormat(CFSTR_SHELLIDLISTOFFSET);
  CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
  CF_FILEDESCRIPTORA := RegisterClipboardFormat(CFSTR_FILEDESCRIPTORA);
  CF_FILEDESCRIPTORW := RegisterClipboardFormat(CFSTR_FILEDESCRIPTORW);
  PIDLMgr := TCommonPIDLManager.Create;

finalization
  PIDLMgr.Free;

end.
