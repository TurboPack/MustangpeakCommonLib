unit MPThreadManager;

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
//
// Special thanks to the following in no particular order for their help/support/code
//    Danijel Malik, Robert Lee, Werner Lehmann, Alexey Torgashin, Milan Vandrovec
//----------------------------------------------------------------------------

interface

{$I Options.inc}
{$I ..\Include\Addins.inc}

{$ifdef COMPILER_12_UP}
  {$WARN IMPLICIT_STRING_CAST       OFF}
 {$WARN IMPLICIT_STRING_CAST_LOSS  OFF}
{$endif COMPILER_12_UP}

{$B-}

// See procedure TEasyThread.ExecuteStub; to understand what this does
{.$DEFINE DEBUG_THREAD}
{.$DEFINE DRAG_OUT_THREAD_SHUTDOWN}


uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Controls,
  ShlObj,
  ShellAPI,
  ActiveX,
  MPShellTypes,
  MPCommonObjects, MPCommonUtilities;


const
  COMMONTHREADFILTERWNDCLASS = 'clsEasyThreadFilter';
  COMMONTHREADSAFETYVALVE = 200;  // Number of PostMessage trys before giving up
  WM_COMMONTHREADCALLBACK = WM_APP + 356;  // Handle this message to recieve the data from the thread
  WM_COMMONTHREADNOTIFIER = WM_APP + 355;  // Used internally to pass the data from the thread to the dispatch window
  TID_START = WPARAM(0);   // Use this Thread ID to start custom ID's for the ThreadRequest.RequestID field
                   // This way the same thread can be used for various tasks and call a common
                   // message handler

  FORCE_KILL_THREAD_COUNT  = 10;    // 100 loops of THREAD_SHUTDOWN_WAIT_DELAY then TerminateThread()
  THREAD_SHUTDOWN_WAIT_DELAY = 200;  // miliseconds

type
  TCommonThreadRequest = class;
  TCommonThreadManager = class;
  TCommonThread = class;
  TPIDLCallbackThreadRequest = class;

  TCommonThreadPriority = 0..100;
  TNamespaceCallbackProc = procedure(Request: TPIDLCallbackThreadRequest) of object;

  // TMessage definition for how the data is passed to the target window via PostMessage
  TWMThreadRequest = {$IFNDEF CPUX64}packed{$ENDIF} record
    Msg: Cardinal;
    {$IFDEF CPUX64}MsgFiller: TDWordFiller;{$ENDIF}
    RequestID: WPARAM;
    Request: TCommonThreadRequest;
    Result: LRESULT;
  end;

  TCommonThreadDirection = (
    etdFirstInFirstOut,     // Requests are serviced from the first to the last
    etdFirstInLastOut       // Requests are serviced from the last to to first
  );

  // **************************************************************************
  // Record that is used to set the name of the Thread.
  // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vsdebug/html/vxtsksettingthreadname.asp
  // http://bdn.borland.com/article/0,1410,29800,00.html
  // **************************************************************************
  TThreadNameInfoA = record
    FType: LongWord;     // must be 0x1000
    FName: PAnsiChar;        // pointer to name (in user address space)
    FThreadID: LongWord; // thread ID (-1 indicates caller thread)
    FFlags: LongWord;    // reserved for future use, must be zero
  end;

  // **************************************************************************
  // Thread Request
  //   Any requests for a TCommonThread to extract data are made through a Thread
  // Request object.  The main thread typically creates the object (or decendent)
  // and fills in the basic info, the item index associated with the request, the
  // Window to notify when finished.
  // The TCommonThreadRequest object pointer is in the lParam of the message.
  // Define the message handle using the TWMRequestThread type for the parameter.
  // i.e.
  //  type
  //    TSomeTWinControl = class(TWinControl)
  //      procedure WMEasyThreadCallback(var Msg: TWMRequestThread); message WM_COMMONTHREADCALLBACK;
  //    end;
  // **************************************************************************
  TCommonThreadRequest = class(TPersistent)
  private
    FID: WPARAM;                     // The ID that identifies the request type
    FPriority: TCommonThreadPriority;  // The Thread will sort the request list by Priority, 0 being highest 100 being the lowest
    FRefCount: Integer;
    FTag: Integer;                   // User defineable field
    FThread: TCommonThread;          // Reference to the thread handling the request
    FWindow: TWinControl;            // The control to send the Message to, set to nil to have the thread free the object without dispatching it to the main thread
    FItem: Pointer;                  // Identifier of the Item the threaded data is being extracted for
    FRemainingRequests: Integer;     // Number of remaining requests in the thread prior to being dispatched to the window
    FCallbackWndMessage: Cardinal;   // This is the window message that is sent to the client window, WM_COMMONTHREADCALLBACK by default
  protected
    property RefCount: Integer read FRefCount write FRefCount;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    function HandleRequest: Boolean; virtual; abstract;
    procedure Assign(Source: TPersistent); override;

    procedure Prioritize(RequestList: TList); virtual;
    procedure Release;
    property CallbackWndMessage: Cardinal read FCallbackWndMessage write FCallbackWndMessage;
    property Item: Pointer read FItem write FItem;
    property ID: WPARAM read FID write FID;
    property Priority: TCommonThreadPriority read FPriority write FPriority default 50;
    property RemainingRequests: Integer read FRemainingRequests write FRemainingRequests;
    property Tag: Integer read FTag write FTag;
    property Thread: TCommonThread read FThread;
    property Window: TWinControl read FWindow write FWindow;
  end;
  TCommonThreadRequestClass = class of TCommonThreadRequest;

  // ***************************************************************
  // A ThreadRequest that uses a PIDL to extract its data from within
  // the context of the thread.
  // ***************************************************************
  TPIDLThreadRequest = class(TCommonThreadRequest)
  private
    FPIDL: PItemIDList;
  public
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    property PIDL: PItemIDList read FPIDL write FPIDL;
  end;

  // ***************************************************************
  // A ThreadRequest that extracts the shell supplied Icon index of an object via a PIDL
  // ***************************************************************
  TShellIconThreadRequest = class(TPIDLThreadRequest)
  private
    FImageIndex: Integer; // [OUT] the image index for the PIDL
    FLarge: Boolean; // [IN] Get the large image index
    FOpen: Boolean;  // [IN] Get the "open" (folder expanded) image index
    FOverlayIndex: Integer;
  public
    function HandleRequest: Boolean; override; // Extracts the Icon from the PIDL
    property ImageIndex: Integer read FImageIndex;
    property Large: Boolean read FLarge write FLarge;
    property Open: Boolean read FOpen write FOpen;
    property OverlayIndex: Integer read FOverlayIndex write FOverlayIndex;
  end;

  // ***************************************************************
  // A ThreadRequest that has a callback function instead of a Window Handle to send a message to
  // ***************************************************************
  TPIDLCallbackThreadRequest = class(TPIDLThreadRequest)
  private
    FCallbackProc: TNamespaceCallbackProc;
    FTargetObject: TObject;
  public
    procedure Assign(Source: TPersistent); override;
    property CallbackProc: TNamespaceCallbackProc read FCallbackProc write FCallbackProc;
    property TargetObject: TObject read FTargetObject write FTargetObject;
  end;

  // **************************************************************************
  // Easy Thread
  //   A thread object that does not use a Syncronize type implementation.  The
  // method of blocking the main thread to service the thread result is ok for
  // a long processing thread that is not used often.  Extracting Icons for an
  // Explorer type listview for instance needs to update the listview very fast
  // and very often.  In this case it is better to use the Window messaging
  // system to post message to the window with the extracted data and allow the
  // window to process it when it can.  This method is much smoother for the GUI
  // but has more challenges to keep it thread safe.  A number of syncronization
  // objects are defined and created to use for various tasks.
  // Use the TCommonThreadManager and its methods to make using TCommonThread easier
  // and safer.
  // **************************************************************************
  TCommonThread = class
  private
    FFreeOnTerminate: Boolean;
    FHandle: THandle;
    FOLEInitialized: Boolean;
    FTargetWnd: HWnd; // Window that the message is posted to.  It will get a WM_COMMONTHREADNOTIFIER message with the TCommonThreadRequest in LParam
    FThreadID: DWORD;
    FStub: ICallBackStub;
    FTerminated: Boolean;
    FSuspended: Boolean;
    FEvent: THandle;
    FCriticalSectionInitialized: Boolean;
    FCriticalSection: TRTLCriticalSection;
    FRefCount: Integer;
    FRequestList: TThreadList;
    FDirection: TCommonThreadDirection;
    FRunning: Boolean;
    FRequestListLocked: Boolean;
    FTempListLock: TList;
    function GetPriority: TThreadPriority;
    procedure SetPriority(const Value: TThreadPriority);
    procedure SetSuspended(const Value: Boolean);
    procedure ExecuteStub;  stdcall;
    function GetLock: TRTLCriticalSection;
    function GetEvent: THandle;
    procedure SetDirection(const Value: TCommonThreadDirection);
    procedure SetRequestListLocked(const Value: Boolean);
  protected
    FFinished: Boolean;
    procedure AddRequest(Request: TCommonThreadRequest; DoSetEvent: Boolean);
    procedure Execute; virtual; abstract; // Called in context of thread
    procedure FinalizeThread; virtual;    // Called in context of thread
    procedure InitializeThread; virtual;  // Called in context of thread

    property CriticalSectionInitialized: Boolean read FCriticalSectionInitialized write FCriticalSectionInitialized;
    property Event: THandle read GetEvent;
    property RequestListLocked: Boolean read FRequestListLocked write SetRequestListLocked; // Don't set/reset this across threads!
    property Stub: ICallbackStub read FStub write FStub;
    property Terminated: Boolean read FTerminated;
  public
    constructor Create(CreateSuspended: Boolean); virtual;
    destructor Destroy; override;

    procedure AddRef;
    procedure FlushRequestList;
    procedure ForceTerminate;
    procedure LockThread;
    procedure Release;
    procedure Resume;
    procedure Terminate; virtual;
    procedure TriggerEvent;
    procedure UnlockThread;
    property Direction: TCommonThreadDirection read FDirection write SetDirection;
    property Finished: Boolean read FFinished;
    property FreeOnTerminate: Boolean read FFreeOnTerminate write FFreeOnTerminate;
    property Handle: THandle read FHandle;
    property Lock: TRTLCriticalSection read GetLock;
    property OLEInitialized: Boolean read FOLEInitialized;
    property Priority: TThreadPriority read GetPriority write SetPriority default tpNormal;
    property RefCount: Integer read FRefCount write FRefCount;
    property RequestList: TThreadList read FRequestList write FRequestList;
    property Running: Boolean read FRunning;
    property Suspended: Boolean read FSuspended write SetSuspended;
    property TargetWnd: HWnd read FTargetWnd write FTargetWnd;
    property ThreadID: DWORD read FThreadID;
  end;
  TCommonBaseThreadClass = class of TCommonThread;

  // **************************************************************************
  // Event Thread
  //   A decendant of TCommonThread that takes the encapsulation a step further.
  // This class defines a thread loop that calls virtual methods that can be
  // overridden in a decendent.  It makes creating a thread VERY easy and
  // very safe.
  // **************************************************************************
  TCommonEventThread = class(TCommonThread)
  private
    FTargetWndNotifyMsg: DWORD;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property TargetWndNotifyMsg: DWORD read FTargetWndNotifyMsg write FTargetWndNotifyMsg;
  end;
  TCommonEventThreadClass = class of TCommonEventThread;

  // **************************************************************************
  // ShellExecute Thread
  //   A decendant of TCommonThread that ShellExecuteEx's in a thread
  // **************************************************************************
  TCommonShellExecuteThread = class(TCommonThread)
  private
    FlpClass: WideString;
    FlpDirectory: WideString;
    FlpFile: WideString;
    FlpParameters: WideString;
    FlpVerb: WideString;
    FPIDL: PItemIDList;
  protected
    procedure Execute; override;
  public
    ShellExecuteInfoA: TShellExecuteInfoA;
    ShellExecuteInfoW: TShellExecuteInfoW;

    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    // Need local variable for the strings and PIDLs so they won't get freed on
    // us before the thread uses them.
    property lpClass: WideString read FlpClass write FlpClass;
    property lpDirectory: WideString read FlpDirectory write FlpDirectory;
    property lpFile: WideString read FlpFile write FlpFile;
    property lpParameters: WideString read FlpParameters write FlpParameters;
    property lpVerb: WideString read FlpVerb write FlpVerb;
    property PIDL: PItemIDList read FPIDL write FPIDL;
  end;


  // **************************************************************************
  // Callback Event Thread
  // **************************************************************************
  TCommonCallbackEventThread = class(TCommonEventThread)
  protected
    procedure Execute; override;
  end;

  // **************************************************************************
  // Thread Manager
  //  A class the encapsulate the TCommonThread or decendant. The Thread filters
  // all requests to a window created in this object for dispatch to the
  // desired window and messageID.  By accessing the Thread through the methods
  // in the class using the thread is simple and safe from race conditions since
  // the class handles the syncronization of the data for you.
  // Simply register a TWinControl decendent and create a TCommonThreadRequest
  // decendant. Make SURE to override the Assign method, the thread must make
  // a copy of each object to use to extract the data.
  // This copy is what is sent to the registered window, not the original
  // that was added to the Request list through AddRequest.
  // **************************************************************************
  TCommonThreadManager = class(TComponent)
  private
    FAClassName: AnsiString;
    FControlList: TThreadList;
    FStub: ICallBackStub;
    FFilterWindow: HWND;
    FEnabled: Boolean;

    function GetThread: TCommonThread;
    function GetFilterWindow: HWND;
    function GetRequestCount: Integer;
    procedure SetEnabled(const Value: Boolean);
  protected
    FThread: TCommonThread;

    function FilterWndProc(Wnd: HWND; uMsg: UINT; wParam: WPARAM; lParam: LPARAM): LRESULT; stdcall;
    procedure CreateThreadObject; virtual;
    function FindControl(Window: TWinControl; LockedList: TList): Integer;
    procedure DispatchRequest(lParam: LPARAM; wParam: WPARAM); virtual;
    procedure FreeThread;
    procedure InternalUnRegisterControl(Window: TWinControl; LockedControlList: TList);
    procedure RegisterFilterWindow;

    property AClassName: AnsiString read FAClassName write FAClassName;
    property ControlList: TThreadList read FControlList write FControlList;
    property FilterWindow: HWND read GetFilterWindow write FFilterWindow;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure AddRequest(Request: TCommonThreadRequest; DoSetEvent: Boolean);
    procedure FlushAllMessageCache(Window: TWinControl; Item: Pointer = nil);
    procedure FlushMessageCache(Window: TWinControl; RequestID: WPARAM; Item: Pointer = nil);
    function RegisterControl(Window: TWinControl): Boolean;
    procedure UnRegisterAll;
    procedure UnRegisterControl(Window: TWinControl);

    property RequestCount: Integer read GetRequestCount;
    property Thread: TCommonThread read GetThread;

  published
    property Enabled: Boolean read FEnabled write SetEnabled default False;
  end;

  // **************************************************************************
  // Callback Thread Manager
  // **************************************************************************
  TCallbackThreadManager = class(TCommonThreadManager)
  protected
    procedure CreateThreadObject; override;
    procedure DispatchRequest(lParam: LPARAM; wParam: WPARAM); override;
  public
    procedure AddRequest(Request: TPIDLCallbackThreadRequest; DoSetEvent: Boolean); reintroduce;
    procedure FlushObjectCache(AnObject: TObject);
  end;


function GlobalThreadManager: TCommonThreadManager;
function GlobalCallbackThreadManager: TCallbackThreadManager;

implementation

uses
  MPResources, MPShellUtilities;

var
  PIDLMgr: TCommonPIDLManager;
  GlobalThread: TCommonThreadManager;
  GlobalCallbackThread: TCallbackThreadManager;
  ThreadsAlive: Integer = 0;
  ThreadRequestsAlive: Integer = 0;

function GlobalThreadManager: TCommonThreadManager;
begin
  if not Assigned(GlobalThread) then
  begin
    GlobalThread := TCommonThreadManager.Create(nil);
    GlobalThread.Enabled := True;
  end;
  Result := GlobalThread
end;

function GlobalCallbackThreadManager: TCallbackThreadManager;
begin
  if not Assigned(GlobalCallbackThread) then
  begin
    GlobalCallbackThread := TCallbackThreadManager.Create(nil);
    GlobalCallbackThread.Enabled := True;
  end;
  Result := GlobalCallbackThread
end;

{ TCommonThreadRequest }
destructor TCommonThreadRequest.Destroy;
begin
  Dec(ThreadRequestsAlive);
  inherited Destroy;
end;


procedure TCommonThreadRequest.Assign(Source: TPersistent);
var
  S: TCommonThreadRequest;
begin
  if Source is TCommonThreadRequest then
  begin
    S := TCommonThreadRequest(Source);
    Window := S.Window;
    ID := S.ID;
    Item := S.Item;
    RemainingRequests := S.RemainingRequests;
    CallbackWndMessage := S.CallbackWndMessage;
    Priority := S.Priority;
    Tag := S.Tag
  end
end;

constructor TCommonThreadRequest.Create;
begin
  inherited;
  FCallbackWndMessage := WM_COMMONTHREADCALLBACK;
  Priority := 50;
  Inc(ThreadRequestsAlive)
end;

procedure TCommonThreadRequest.Prioritize(RequestList: TList);
begin
  // Override to allow control of the order the thread requests are sorted
end;

procedure TCommonThreadRequest.Release;
begin
  InterlockedDecrement(FRefCount);
  if RefCount <= 0 then
  begin
    RefCount := 0;
    Free
  end
end;

{ TEasyPIDLThreadRequest }

destructor TPIDLThreadRequest.Destroy;
begin
  PIDLMgr.FreePIDL(PIDL);
  inherited Destroy;
end;

procedure TPIDLThreadRequest.Assign(Source: TPersistent);
begin
  inherited Assign(Source);
  if Source is TPIDLThreadRequest then
    PIDL := PIDLMgr.CopyPIDL(TPIDLThreadRequest( Source).PIDL);
end;

{ TEasyIconThreadRequest }

function TShellIconThreadRequest.HandleRequest: Boolean;

{  function GetIconByIShellIcon(PIDL: PItemIDList; var Index: integer): Boolean;
  var
    Flags: Longword;
    OldCB: Word;
    Old_ID: PItemIDList;
    Desktop, Folder: IShellFolder;
    ShellIcon: IShellIcon;
    OverlayInterface: IShellIconOverlay;
  begin
    Result := False;
    Overlay := -1;
    Index := -1;
    PIDLMgr.StripLastID(PIDL, OldCB, Old_ID);
    try
      SHGetDesktopFolder(Desktop);
      Desktop.BindToObject(PIDL, nil, IShellFolder, Pointer(Folder));
      Old_ID.mkid.cb := OldCB;
      if Assigned(Folder) then
      begin
        if Folder.QueryInterface(IShellIcon, ShellIcon) = S_OK then
        begin
          Flags := GIL_FORSHELL;
          if Open then
            Flags := Flags or GIL_OPENICON;
          Result := ShellIcon.GetIconOf(Old_ID, Flags, Index) = NOERROR
        end;
      end
    finally
      Old_ID.mkid.cb := OldCB
    end
  end;

  procedure GetIconBySHGetFileInfo(APIDL: PItemIDList; var Index: Integer);
  var
    Flags: integer;
    Info: TSHFILEINFOA;
  begin
    Flags := SHGFI_PIDL or SHGFI_SYSICONINDEX or SHGFI_SHELLICONSIZE;
    if Large then
      Flags := Flags or SHGFI_LARGEICON
    else
      Flags := Flags or SHGFI_SMALLICON;

    if Open then
      Flags := Flags or SHGFI_OPENICON;
      
    if SHGetFileInfoA(PAnsiChar(APIDL), 0, Info, SizeOf(Info), Flags) <> 0 then
      Index := Info.iIcon
    else
      Index := 0
  end;    }

var
  NS: TNamespace;
begin
{  Result := GetIconByIShellIcon(PIDL, FImageIndex);
  if not Result then
    GetIconBySHGetFileInfo(PIDL, FImageIndex);   }
  NS := TNamespace.Create(PIDL, nil);
  NS.FreePIDLOnDestroy := False;
  if Large then
    FImageIndex := NS.GetIconIndex(Open, icLarge, True)
  else
    FImageIndex := NS.GetIconIndex(Open, icSmall, True);
  FOverlayIndex := NS.OverlayIndex;
  NS.Free;
  Result := True;
end;

{ TPIDLCallbackThreadRequest }

procedure TPIDLCallbackThreadRequest.Assign(Source: TPersistent);
begin
  inherited Assign(Source);
  CallbackProc := (Source as TPIDLCallbackThreadRequest).CallbackProc;
  TargetObject := (Source as TPIDLCallbackThreadRequest).TargetObject
end;

{ TCommonThread }

procedure TCommonThread.AddRef;
begin
  InterlockedIncrement(FRefCount);
end;

procedure TCommonThread.AddRequest(Request: TCommonThreadRequest; DoSetEvent: Boolean);
var
  List: TList;
begin
  List := RequestList.LockList;
  try
    List.Add(Request);
    if DoSetEvent then
      TriggerEvent;
  finally
    RequestList.UnlockList
  end
end;

constructor TCommonThread.Create(CreateSuspended: Boolean);
var
  Flags: DWORD;
begin
  Inc(ThreadsAlive);
  IsMultiThread := True;
  Direction := etdFirstInLastOut;
  RequestList := TThreadList.Create;
  Stub := TCallBackStub.Create(Self, @TCommonThread.ExecuteStub, 0);
  Flags := 0;
  if CreateSuspended then
  begin
    Flags := CREATE_SUSPENDED;
    FSuspended := True
  end;
  FHandle := CreateThread(nil, 0, Stub.StubPointer, nil, Flags, FThreadID);
end;

destructor TCommonThread.Destroy;
var
  i: Integer;
  List: TList;
begin
  Assert(Finished, 'The Thread must be terminated before destroying the TCommonThread object');
  if Handle <> 0 then
    CloseHandle(Handle);
  FHandle := 0;
  if Event <> 0 then
    CloseHandle(Event);
  FEvent := 0;
  if CriticalSectionInitialized then
    DeleteCriticalSection(FCriticalSection);
  List := RequestList.LockList;
  try
    for i := 0 to List.Count - 1 do
      TObject(List[i]).Free;
    List.Count := 0;
  finally
    RequestList.UnlockList;
  end;
  FreeAndNil(FRequestList);
  FRequestList := nil;
  Dec(ThreadsAlive);
  inherited;
end;

procedure TCommonThread.ExecuteStub;
// Called in the context of the thread
{$IFDEF DEBUG_THREAD}
var
  ThreadNameInfoA: TThreadNameInfoA;
{$ENDIF}
begin
  {$IFDEF DEBUG_THREAD}
  if IsWinNT then
  begin
    // Set the name for the thread to debug it with the Thread View panel.
    // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vsdebug/html/vxtsksettingthreadname.asp
    // http://bdn.borland.com/article/0,1410,29800,00.html
    ThreadNameInfoA.FType := $1000;
    ThreadNameInfoA.FName := PAnsiChar(AnsiString(ClassName));
    ThreadNameInfoA.FThreadID := $FFFFFFFF;
    ThreadNameInfoA.FFlags := 0;
    try
      RaiseException($406D1388, 0, sizeof(ThreadNameInfoA) div sizeof(LongWord), @ThreadNameInfoA);
    except
    end;
  end;
  {$ENDIF}

  try
    FRunning := True;
    InitializeThread;
    try
      Execute
    except
    end
  finally
    FinalizeThread;
    {$IFDEF DRAG_OUT_THREAD_SHUTDOWN}
    Sleep(3000);
    {$ENDIF}
    if FreeOnTerminate then
    begin
      // If FreeOnTerminate then the user can't expect to look at these
      // variables since they can't be sure when the object will be freed
      FRunning := False;
      FFinished := True;
      FHandle := 0;
      Free;
      ExitThread(0);
    end else
    begin
      // Set these here or there will be a race condtion as to when the
      // main thread frees the object and we still need to access
      // local variables (like FreeOnTerminate)
      FRunning := False;
      FFinished := True;
      FHandle := 0;
      ExitThread(0);
    end
  end
end;

function TCommonThread.GetEvent: THandle;
begin
  if FEvent = 0 then
    FEvent := CreateEvent(nil, True, False, nil);
  Result := FEvent;
end;

function TCommonThread.GetLock: TRTLCriticalSection;
begin
  if not CriticalSectionInitialized then
  begin
    InitializeCriticalSection(FCriticalSection);
    CriticalSectionInitialized := True
  end;
  Result := FCriticalSection
end;

function TCommonThread.GetPriority: TThreadPriority;
var
  P: Integer;
begin
  Result := tpNormal;
  P := GetThreadPriority(FHandle);
  case P of
    THREAD_PRIORITY_IDLE:          Result := tpIdle;
    THREAD_PRIORITY_LOWEST:        Result := tpLowest;
    THREAD_PRIORITY_BELOW_NORMAL:  Result := tpLower;
    THREAD_PRIORITY_NORMAL:        Result := tpNormal;
    THREAD_PRIORITY_HIGHEST:       Result := tpHigher;
    THREAD_PRIORITY_ABOVE_NORMAL:  Result := tpHighest;
    THREAD_PRIORITY_TIME_CRITICAL: Result := tpTimeCritical;
  end
end;

procedure TCommonThread.FinalizeThread;
begin
  try
    if OLEInitialized then
      OLEUnInitialize
  except
  end
end;

procedure TCommonThread.FlushRequestList;
var
  List: TList;
  i: Integer;
  Request: TObject;
begin
  List := RequestList.LockList;
  try
    for i := List.Count - 1 downto 0 do
    begin
      Request := TObject(TObject(List[i]));
      List.Delete(i);
      Request.Free;
    end
  finally
    RequestList.UnlockList
  end
end;

procedure TCommonThread.InitializeThread;
begin
  FOLEInitialized := Succeeded( OLEInitialize(nil))
end;

procedure TCommonThread.LockThread;
begin
  if not CriticalSectionInitialized then
    Lock;
  EnterCriticalSection(FCriticalSection)
end;

procedure TCommonThread.Release;
begin
  InterlockedDecrement(FRefCount);
end;

procedure TCommonThread.Resume;
begin
  Suspended := False
end;

procedure TCommonThread.SetDirection(const Value: TCommonThreadDirection);
begin
  FDirection := Value;
end;

procedure TCommonThread.SetPriority(const Value: TThreadPriority);
begin
  case Value of
    tpIdle        : SetThreadPriority(Handle,  THREAD_PRIORITY_IDLE);
    tpLowest      : SetThreadPriority(Handle, THREAD_PRIORITY_LOWEST);
    tpLower       : SetThreadPriority(Handle, THREAD_PRIORITY_BELOW_NORMAL);
    tpNormal      : SetThreadPriority(Handle, THREAD_PRIORITY_NORMAL);
    tpHigher      : SetThreadPriority(Handle, THREAD_PRIORITY_HIGHEST);
    tpHighest     : SetThreadPriority(Handle, THREAD_PRIORITY_ABOVE_NORMAL);
    tpTimeCritical: SetThreadPriority (Handle, THREAD_PRIORITY_TIME_CRITICAL);
  end
end;

procedure TCommonThread.SetRequestListLocked(const Value: Boolean);
begin
  if FRequestListLocked <> Value then
  begin
    if Value then
      FTempListLock := RequestList.LockList
    else
      RequestList.UnlockList;
    FRequestListLocked := Value;
  end
end;

procedure TCommonThread.SetSuspended(const Value: Boolean);
begin
  if FSuspended <> Value then
  begin
    if Handle <> 0 then
    begin
      if Value then
        SuspendThread(FHandle)
      else
        ResumeThread(FHandle);
      FSuspended := Value;
    end
  end
end;

procedure TCommonThread.Terminate;
begin
  Suspended := False;
  FTerminated := True;
  TriggerEvent;
end;

procedure TCommonThread.TriggerEvent;
begin
  SetEvent(Event);
end;

procedure TCommonThread.ForceTerminate;
var
  Temp: THandle;
begin
  Temp := Handle;
  if Temp <> 0 then
  begin
    FHandle := 0;
    FRunning := False;
    FFinished := True;
    TerminateThread(Temp, 0);
  end
end;

procedure TCommonThread.UnlockThread;
begin
  if CriticalSectionInitialized then
    LeaveCriticalSection(FCriticalSection)
end;

{ TCommonEventThread }

function PrioritizeSort(Item1, Item2: Pointer): Integer;
begin
  Result := TCommonThreadRequest(Item1).Priority - TCommonThreadRequest(Item2).Priority
end;

constructor TCommonEventThread.Create(CreateSuspended: Boolean);
begin
  inherited Create(CreateSuspended);
  TargetWndNotifyMsg := WM_COMMONTHREADNOTIFIER;  // Default for the internal dispatch Window
  // Create the event
  Event;
end;

destructor TCommonEventThread.Destroy;
begin
  inherited Destroy;
end;

procedure TCommonEventThread.Execute;
var
  ARequestCopy, OriginalRequest: TCommonThreadRequest;
  List: TList;
  WorkingIndex, SafetyValve: Integer;
  LoopCount: Cardinal;
  RequestClassType: TCommonThreadRequestClass;
begin
  LoopCount := 0;
  while not Terminated and (TargetWnd <> 0) do
  try
    WaitForSingleObject(Event, INFINITE);
    if not Terminated then
    begin
      // Small breather so thread does not steal too much processor time
      Sleep(0);
      // Take the Request from the list and make a copy of it to work with.
      // This is in case a control purges the list of requests, this way the
      // thread is working on a copy and when it it ready it sees if the original
      // is still in the list.  If not it dumps it.
      List := RequestList.LockList;
      try
        // Only do this every so many extracts.  It can be rather slow with lots of items
        if LoopCount mod 10 = 0 then
          List.Sort(PrioritizeSort);

        ARequestCopy := nil;
        OriginalRequest := nil;

        if List.Count > 0 then
        begin
          if Direction = etdFirstInFirstOut then
            WorkingIndex := 0
          else
            WorkingIndex := List.Count - 1;
          // Make a copy of the pointer to the original Request
          OriginalRequest := TCommonThreadRequest(List[WorkingIndex]);
          // Need to make a copy of the Request.  This is because while we are
          // extracting the data for the Request the main thread may do something
          // to the list that renders the object in the list invalid.  By working
          // with a copy we do not worry about it.  We will later check to make
          // sure that the object on the list is still valid before dispatching the
          // results to the main thread

          RequestClassType := TCommonThreadRequestClass(OriginalRequest.ClassType);
          ARequestCopy := RequestClassType.Create;
          ARequestCopy.Assign(OriginalRequest);
        end else
        begin
          WorkingIndex := -1;
          // Reset the event to enter the WaitForSingleObject method again
          ResetEvent(FEvent)
        end
      finally
        RequestList.UnlockList
      end;

      if Assigned(ARequestCopy) then
      begin
        ARequestCopy.FThread := Self;
        // Extract the data for the Request
        if ARequestCopy.HandleRequest then
        begin
          List := RequestList.LockList;
          try
            // Check to see if the WorkingIndex is still valid in the list.
            // If not the item will be left if the queue and it will have to be done again
            if (List.Count > 0) and (WorkingIndex < List.Count) then
            begin
              // Now make sure the actual object is still in the list in the same
              // position.  If not then we will leave it in the list and extract it
              // again.
              if (OriginalRequest = TCommonThreadRequest(List[WorkingIndex])) then
              begin
                // It still exists so we can delete it as we will dispatch it
                List.Delete(WorkingIndex);
                ARequestCopy.RemainingRequests := List.Count;
                // Try to Post it to the main thread through the ThreadManagers window
                SafetyValve := 0;
                while not PostMessage(TargetWnd, TargetWndNotifyMsg, 0,
                  LPARAM(ARequestCopy)) and (SafetyValve < COMMONTHREADSAFETYVALVE) do
                begin
                  Inc(SafetyValve);
                  Sleep(10);
                end;

                // If failed (VERY, VERY, VERY unlikely) then add it back to the list
                // and try it again later
                if SafetyValve >= COMMONTHREADSAFETYVALVE then
                begin
                  List.Add(OriginalRequest);
                  FreeAndNil(ARequestCopy)
                end else
                  FreeAndNil(OriginalRequest);
              end else
                FreeAndNil(ARequestCopy)
            end else
              FreeAndNil(ARequestCopy)
          finally
            RequestList.UnlockList
          end
        end else
          FreeAndNil(ARequestCopy)
      end;
      Inc(LoopCount)
    end
  except
  end
end;

{ TCommonThreadManager }

procedure TCommonThreadManager.AddRequest(Request: TCommonThreadRequest; DoSetEvent: Boolean);
var
  DoAdd: Boolean;
begin
  if Assigned(Request) and Enabled then
  begin
    DoAdd := False;
    if (Request is TPIDLCallbackThreadRequest) then
      DoAdd := Assigned( TPIDLCallbackThreadRequest(Request).CallbackProc) and
        Assigned(TPIDLCallbackThreadRequest(Request).TargetObject);
    if not DoAdd then
      DoAdd := (FindControl(Request.Window, nil) > -1) or Assigned(Request.Window);
    Assert(DoAdd, STR_UNREGISTEREDCONTROL);
    if DoAdd and Assigned(Thread) then
      Thread.AddRequest(Request, DoSetEvent);
  end else
    Request.Free
end;

constructor TCommonThreadManager.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  ControlList := TThreadList.Create;
  AClassName := COMMONTHREADFILTERWNDCLASS + IntToStr( Integer( Self))
end;

procedure TCommonThreadManager.CreateThreadObject;
begin
  FThread := TCommonEventThread.Create(True);
  FThread.TargetWnd := FilterWindow;
  Thread.Resume;
end;

destructor TCommonThreadManager.Destroy;
begin
  UnRegisterAll;
  FlushMessageCache(nil, TID_START);
  // Thread Freed with last client is unregistered
  // Safe to destroy the window now
  if FFilterWindow <> 0 then
    DestroyWindow(FFilterWindow);
  FFilterWindow := 0;
  // Unregister the window class.  If another thread is also using this class
  // windows will not unregister it until the last thread has destroyed any windows
  // based on this class
  if AClassName <> '' then
    Windows.UnregisterClassA(PAnsiChar( AnsiString(AClassName)), hInstance);
  FreeAndNil(FControlList);
  FreeThread;
  inherited;
end;

function TCommonThreadManager.FilterWndProc(Wnd: HWND; uMsg: UINT;
  wParam: WPARAM; lParam: LPARAM): LRESULT;
begin
  Result := 0;
  case uMsg of
    WM_NCCREATE:  Result := 1;
    WM_COMMONTHREADNOTIFIER: DispatchRequest(lParam, wParam);
  else
    Result := DefWindowProc(Wnd, uMsg, wParam, lParam);
  end
end;

function TCommonThreadManager.FindControl(Window: TWinControl; LockedList: TList): Integer;
//
// Loops the Window/MessageID pairs in the ControlList looking for a match
//
var
  List: TList;
  I: Integer;
  Found: Boolean;
begin
  Result := -1;
  if not Assigned(LockedList) then
    List := ControlList.LockList
  else
    List := LockedList;

  try
    I := 0;
    Found := False;
    while (I < List.Count) and not Found do
    begin
      if (Window = TWinControl(List[I])) then
      begin
        Result := I;
        Found := True
      end;
      Inc(i);
    end
  finally
    if not Assigned(LockedList) then
      ControlList.UnlockList
  end
end;

procedure TCommonThreadManager.DispatchRequest(lParam: LPARAM; wParam: WPARAM);
var
  i: Integer;
  RegList: TList;
  Request: TCommonThreadRequest;
  RequestList: TList;
begin
  Request := TCommonThreadRequest(lParam);
  try
    RequestList := Thread.RequestList.LockList;
    try
      // Allow the Request to prioritize the list if it wants to
      Request.Prioritize(RequestList);

      if not Assigned(Request.Window) then
      begin
        RegList := ControlList.LockList;
        try
          Request.RefCount := 1;
          for i := 0 to RegList.Count - 1 do
          begin
            if TWinControl(RegList[i]).HandleAllocated then
              Request.RefCount := RegList.Count + 1;
          end;
                // It is a broadcast message
          wParam := Request.ID;
          for i := 0 to RegList.Count - 1 do
          begin
            if TWinControl(RegList[i]).HandleAllocated then
              SendMessage(TWinControl(RegList[i]).Handle, Request.CallbackWndMessage, wParam, lParam);
          end
        finally
          ControlList.UnlockList
        end
      end;

          // Not a broadcast message
        if Assigned(Request.Window) then
        begin
          Request.RefCount := 2;
          if Request.Window.HandleAllocated then
          begin
            wParam := Request.ID;
            SendMessage(Request.Window.Handle, Request.CallbackWndMessage, wParam, lParam);
          end
        end;
    finally
      Thread.RequestList.UnLockList;
    end;
  finally
    Request.Release
  end;
end;

procedure TCommonThreadManager.FlushAllMessageCache(Window: TWinControl; Item: Pointer = nil);
// Flushes all the message cache
begin
  FlushMessageCache(Window, TID_START, Item);
end;

procedure TCommonThreadManager.FlushMessageCache(Window: TWinControl; RequestID: WPARAM; Item: Pointer = nil);
// First locks the thread by locking its RequestList.  This stops the thread
// from accessing a new request.  It then flushes the Windows message cache
// of pending messages matching the RequestId.
// Next any pending requests in the Threads RequestList matching the RequestID
// are removed and freed.
// If Item <> nil then the method will only remove requests for that particular Item
// (comparing it to the TCommonThreadRequest.Item field)
// Pass a RequestID = TID_START to remove all request from a window
var
  Msg: TMsg;
  List: TList;
  I: Integer;
  R: TCommonThreadRequest;
  RepostQuitMsg: Boolean;
  QuitMsgExitCode: Integer;
begin
  List := nil;

  if Enabled then
  begin
    RepostQuitMsg := False;
    QuitMsgExitCode := 0;
    // If the thread is not created yet then there is no point in flushing it
    if Assigned(FThread) then
      List := Thread.RequestList.LockList;
    try
      // Remove the requests in the hidden dispatch message cache
      // I have seen PeekMessage return true and return the WM_QUIT message, this
      // strips the queue of the message to shut down the app!  Don't strip it out
      // until we check to see if it is the right message
      while PeekMessage(Msg, FilterWindow, WM_COMMONTHREADNOTIFIER, WM_COMMONTHREADNOTIFIER, PM_REMOVE) do
      begin
        if Msg.Message = WM_QUIT then
        begin
          QuitMsgExitCode := Msg.wParam;
          RepostQuitMsg := True;
        end else
        begin
          // If the message is for the window to flush then free it, else dispatch it normally
          R := TCommonThreadRequest(Msg.lParam);

          if (R.Window = Window) and ((RequestID = TID_START) or (R.ID = RequestID)) then
          begin
            if Item <> nil then
            begin
              // If is the target item then remove it, else put it back in the queue
              // WARNING don't SendMessage as that could cause some nasty reentrant isses
              if R.Item = Item then
                R.Release
              else
                PostMessage(R.Window.Handle, R.CallbackWndMessage, Msg.wParam, Msg.lParam)
            end else
              R.Release
          end else
          begin
            if R.Window.HandleAllocated then
              SendMessage(R.Window.Handle, R.CallbackWndMessage, R.ID, Msg.lParam)
         //     SendMessage(R.Window.Handle, R.CallbackWndMessage, Msg.wParam, Msg.lParam)
            else
              R.Release
          end
        end
      end;

      if Assigned(Window) then
      begin
        // Remove the requests in the windows cache
        if Window.HandleAllocated then
        begin
          // I have seen PeekMessage return true and return the WM_QUIT message, this
          // strips the queue of the message to shut down the app!  Don't strip it out
          // until we check to see if it is the right message
          while PeekMessage(Msg, Window.Handle, WM_COMMONTHREADCALLBACK, WM_COMMONTHREADCALLBACK, PM_REMOVE) do
          begin
            if Msg.Message = WM_QUIT then
            begin
              QuitMsgExitCode := Msg.wParam;
              RepostQuitMsg := True;
            end else
            begin
              // If the message is for the window to flush then free it
              R := TCommonThreadRequest(Msg.lParam);

              if (R.Window = Window) and ((RequestID = TID_START) or (R.ID = RequestID)) then
              begin
                if Item <> nil then
                begin
                  // If is the target item then remove it, else put it back in the queue
                  // WARNING don't SendMessage as that could cause some nasty reentrant isses
                  if R.Item = Item then
                    R.Release
                  else
                    PostMessage(R.Window.Handle, R.CallbackWndMessage, Msg.wParam, Msg.lParam)
                end else
                  R.Release
              end
            end
          end;
          if RepostQuitMsg then
            PostQuitMessage(QuitMsgExitCode);
        end
      end;

      if Assigned(List) then
      begin
        // Now remove any waiting requests from the list in the thread
        for I := List.Count - 1 downto 0 do
        begin
          R := TCommonThreadRequest(List[I]);
          if (Window = nil) or ((R.Window = Window) and ((RequestID = TID_START) or (R.ID = RequestID))) then
          begin
            if Item <> nil then
            begin
              // Only remove the target item if assigned
              if Item = R.Item then
              begin
                R.Release;
                List.Delete(I);
              end
            end else
            begin
              R.Release;
              List.Delete(I);
            end
          end
        end;
      end
    finally
      if Assigned(FThread) then
        Thread.RequestList.UnlockList
    end;
  end;
end;

function TCommonThreadManager.GetFilterWindow: HWND;
begin
  if FFilterWindow = 0 then
  begin
    RegisterFilterWindow;
  end;
  Result := FFilterWindow;
end;

function TCommonThreadManager.GetRequestCount: Integer;
var
  List: TList;
begin
  if Enabled then
  begin
    List := Thread.RequestList.LockList;
    try
      Result := List.Count
    finally
      Thread.RequestList.UnlockList
    end
  end else
    Result := 0
end;

function TCommonThreadManager.GetThread: TCommonThread;
begin
  if not Assigned(FThread) then
    CreateThreadObject;
  Result := FThread;
end;

procedure TCommonThreadManager.FreeThread;
var
  i: Integer;
begin        
  if Assigned(FThread) then
  begin
    try
      i := 0;
      // Signal the thread the terminate
      FThread.Terminate;
      while not FThread.Finished do
      begin
        Sleep(THREAD_SHUTDOWN_WAIT_DELAY);
        if i > FORCE_KILL_THREAD_COUNT then
        begin
          FThread.ForceTerminate;
          Break
        end;
        Inc(i)
      end;
    finally
      // Done with the thread
      FreeAndNil(FThread);
    end
  end
end;

procedure TCommonThreadManager.InternalUnRegisterControl(Window: TWinControl; LockedControlList: TList);
//
// Unregisters the Window/MessageID pair
//
var
  List: TList;
  I: Integer;
begin
  if not Assigned(LockedControlList) then
    List := ControlList.LockList
  else
    List := LockedControlList;

  try
    if Enabled then
    begin
      // Lock the Thread from accessing its Request List until we are finished
      // If the thread is not created yet there is no point creating it
      if Assigned(FThread) then
        Thread.RequestListLocked := True;
      try
        I := FindControl(Window, List);
        if I > -1 then
        begin
          FlushAllMessageCache(Window);
          List.Delete(I);
        end;
      finally
        if Assigned(FThread) then
          Thread.RequestListLocked := False;
      end
    end else
    begin
      I := FindControl(Window, List);
      if I > -1 then
      begin
        FlushAllMessageCache(Window);
        List.Delete(i);
      end
    end
  finally
    if not Assigned(LockedControlList) then
      ControlList.UnlockList
  end

end;

function TCommonThreadManager.RegisterControl(Window: TWinControl): Boolean;
var
  List: TList;
begin
  RegisterFilterWindow;
  List := ControlList.LockList;
  try
    Result := List.Add(Window) > -1
  finally
    ControlList.UnlockList
  end
end;

procedure TCommonThreadManager.RegisterFilterWindow;
var
  ClassInfo: TWndClassA;
begin
  if FFilterWindow = 0 then
  begin
    if not GetClassInfoA(hInstance, PAnsiChar( AnsiString(AClassName)), ClassInfo) then
    begin
      if not Assigned(FStub) then
        FStub := TCallBackStub.Create(Self, @TCommonThreadManager.FilterWndProc, 4);
      ClassInfo.style := 0;
      ClassInfo.lpfnWndProc := FStub.StubPointer;
      ClassInfo.cbClsExtra := 0;
      ClassInfo.cbWndExtra := 0;
      ClassInfo.hInstance := hInstance;
      ClassInfo.hIcon := 0;
      ClassInfo.hCursor := 0;
      ClassInfo.hbrBackground := 0;
      ClassInfo.lpszMenuName := '';
      ClassInfo.lpszClassName := PAnsiChar( AnsiString(AClassName));
      Windows.RegisterClassA(ClassInfo);
    end;
    FFilterWindow := CreateWindowA(PAnsiChar( AClassName), '', 0, 0, 0, 0, 0, 0, 0, hInstance, nil);
  end
end;

procedure TCommonThreadManager.SetEnabled(const Value: Boolean);
begin
  if Value <> FEnabled then
  begin
    if not (csDesigning in ComponentState) then
    begin
      if not Value then
      begin
        if Assigned(Thread) then
        begin
          if not Thread.Terminated then
          begin
            Thread.Terminate;
            while not Thread.Finished do
              Sleep(200);
          end;
          FreeThread;
        end
      end
    end;
    FEnabled := Value
  end
end;

procedure TCommonThreadManager.UnRegisterAll;
var
  List: TList;
  i: Integer;
begin
  List := ControlList.LockList;
  try
    for i := List.Count - 1 downto 0 do
      InternalUnRegisterControl(TWinControl(List[i]), List);
  finally
    ControlList.UnlockList
  end
end;

procedure TCommonThreadManager.UnRegisterControl(Window: TWinControl);
var
  List: TList;
begin
  List := ControlList.LockList;
  try
    InternalUnRegisterControl(Window, List)
  finally
    ControlList.UnlockList;
  end;
end;

procedure TCallbackThreadManager.AddRequest(Request: TPIDLCallbackThreadRequest;
  DoSetEvent: Boolean);
begin
  inherited AddRequest(Request, DoSetEvent)
end;

procedure TCallbackThreadManager.CreateThreadObject;
begin
  FThread := TCommonCallbackEventThread.Create(True);
  FThread.TargetWnd := FilterWindow;
  Thread.Resume;
  RegisterFilterWindow;
end;

procedure TCallbackThreadManager.DispatchRequest(lParam: LPARAM; wParam: WPARAM);
var
  RequestList: TList;
  Request: TPIDLCallbackThreadRequest;
begin
  Request := TPIDLCallbackThreadRequest(lParam);
  try
    Request.FRefCount := 1;
    RequestList := Thread.RequestList.LockList;
    try
      // Allow the Request to prioritize the list if it wants to
      Request.Prioritize(RequestList);
      if Assigned(Request.CallbackProc) then
      begin
        Request.FRefCount := 2;
        Request.CallbackProc(Request)
      end
    finally
      Thread.RequestList.UnLockList;
    end;
  finally
    Request.Release
  end
end;

procedure TCallbackThreadManager.FlushObjectCache(AnObject: TObject);
var
  List: TList;
  i: Integer;
begin
  if Assigned(Thread) then
  begin
    List := Thread.RequestList.LockList;
    try
      for i := 0 to List.Count - 1 do
      begin
        if TPIDLCallbackThreadRequest( List[i]).TargetObject = AnObject then
        begin
          List.Delete(i);
          Exit;
        end
      end;
    finally
      Thread.RequestList.UnlockList
    end
  end
end;

{ TCommonCallbackEventThread }
procedure TCommonCallbackEventThread.Execute;
begin
  inherited Execute;
end;

{ TCommonShellExecuteThread }
constructor TCommonShellExecuteThread.Create(CreateSuspended: Boolean);
begin
  inherited;
  FreeOnTerminate := True;
  FillChar(ShellExecuteInfoW, SizeOf(ShellExecuteInfoW), #0);
  FillChar(ShellExecuteInfoA, SizeOf(ShellExecuteInfoA), #0);
end;

destructor TCommonShellExecuteThread.Destroy;
begin
  PIDLMgr.FreePIDL(FPIDL);
  inherited Destroy;
end;

procedure TCommonShellExecuteThread.Execute;
begin
  if IsUnicode then
  begin
    if lpClass <> '' then
      ShellExecuteInfoW.lpClass := PWideChar(lpClass);
    if lpDirectory <> '' then
      ShellExecuteInfoW.lpDirectory := PWideChar(lpDirectory);
    if lpFile <> '' then
      ShellExecuteInfoW.lpFile := PWideChar(lpFile);
    if lpParameters <> '' then
      ShellExecuteInfoW.lpParameters := PWideChar(lpParameters);
    if lpVerb <> '' then
      ShellExecuteInfoW.lpVerb := PWideChar(lpVerb);
    ShellExecuteInfoW.lpIDList := PIDL;
    ShellExecuteExW_MP(@ShellExecuteInfoW);
  end else
  begin
    if lpClass <> '' then
      ShellExecuteInfoA.lpClass := PAnsiChar(AnsiString( lpClass));
    if lpDirectory <> '' then
      ShellExecuteInfoA.lpDirectory := PAnsiChar(AnsiString( lpDirectory));
    if lpFile <> '' then
      ShellExecuteInfoA.lpFile := PAnsiChar(AnsiString( lpFile));
    if lpParameters <> '' then
      ShellExecuteInfoA.lpParameters := PAnsiChar(AnsiString( lpParameters));
    if lpVerb <> '' then
      ShellExecuteInfoA.lpVerb := PAnsiChar(AnsiString( lpVerb));
    ShellExecuteInfoA.lpIDList := PIDL;
    ShellExecuteExA(@ShellExecuteInfoA);
  end
end;

initialization
  IsMultiThread := True;
  PIDLMgr := TCommonPIDLManager.Create;


finalization
  FreeAndNil(GlobalThread);
  FreeAndNil(GlobalCallbackThread);
  FreeAndNil(PIDLMgr);

end.
