(**************************************************************************)
(*                                                                        *)
(* Module:  Unit 'DynamicDLL'       Copyright (c) 1997                    *)
(*                                                                        *)
(* Version: 3.0                     Dr. Dietmar Budelsky                  *)
(* Sub-Version: 0.33                dbudelsky@web.de                      *)
(*                                  Germany                               *)
(*                                                                        *)
(*                                  Morgan Martinet                       *)
(*                                  4723 rue Brebeuf                      *)
(*                                  H2J 3L2 MONTREAL (QC)                 *)
(*                                  CANADA                                *)
(*                                  e-mail: p4d@mmm-experts.com           *)
(*                                                                        *)
(*  look at the project page at: http://python4Delphi.googlecode.com/     *)
(**************************************************************************)
(*  Functionality:  Delphi Components that provide an interface to the    *)
(*                  Python language (see python.txt for more infos on     *)
(*                  Python itself).                                       *)
(*                                                                        *)
(**************************************************************************)
(*  Contributors:                                                         *)
(*      Grzegorz Makarewicz (mak@mikroplan.com.pl)                        *)
(*      Andrew Robinson (andy@hps1.demon.co.uk)                           *)
(*      Mark Watts(mark_watts@hotmail.com)                                *)
(*      Olivier Deckmyn (olivier.deckmyn@mail.dotcom.fr)                  *)
(*      Sigve Tjora (public@tjora.no)                                     *)
(*      Mark Derricutt (mark@talios.com)                                  *)
(*      Igor E. Poteryaev (jah@mail.ru)                                   *)
(*      Yuri Filimonov (fil65@mail.ru)                                    *)
(*      Stefan Hoffmeister (Stefan.Hoffmeister@Econos.de)                 *)
(*      Michiel du Toit (micdutoit@hsbfn.com) - Lazarus Port              *)
(*      Chris Nicolai (nicolaitanes@gmail.com)                            *)
(*      Kiriakos Vlahos (pyscripter@gmail.com)                            *)
(*      Andrey Gruzdev      (andrey.gruzdev@gmail.com)                    *)
(**************************************************************************)
(* This source code is distributed with no WARRANTY, for no reason or use.*)
(* Everyone is allowed to use and change this code free for his own tasks *)
(* and projects, as long as this header and its copyright text is intact. *)
(* For changed versions of this code, which are public distributed the    *)
(* following additional conditions have to be fullfilled:                 *)
(* 1) The header has to contain a comment on the change and the author of *)
(*    it.                                                                 *)
(* 2) A copy of the changed source has to be sent to the above E-Mail     *)
(*    address or my then valid address, if this is possible to the        *)
(*    author.                                                             *)
(* The second condition has the target to maintain an up to date central  *)
(* version of the component. If this condition is not acceptable for      *)
(* confidential or legal reasons, everyone is free to derive a component  *)
(* or to generate a diff file to my or other original sources.            *)
(* Dr. Dietmar Budelsky, 1997-11-17                                       *)
(**************************************************************************)

{$I Definition.Inc}

unit DynamicDll;

{ TODO -oMMM : implement tp_as_buffer slot }
{ TODO -oMMM : implement Attribute descriptor and subclassing stuff }

{$IFNDEF FPC}
  {$IFNDEF DELPHI2010_OR_HIGHER}
      Error! Delphi 7 or higher is required!
  {$ENDIF}
{$ENDIF}

interface

uses
  Types,
{$IFDEF MSWINDOWS}
  Windows,
{$ELSE}
{$IFDEF FPC}
  dynlibs,
{$ELSE}
{$IFDEF LINUX}
  Libc,
{$ENDIF}
{$ENDIF}
{$ENDIF}
  Classes,
  SysUtils,
  SyncObjs,
  Variants,
  WideStrings,
  MethodCallBack;

type
  EDLLImportError = class(Exception)
    public
      WrongFunc : string;
      ErrorCode : Integer;
  end;

//-------------------------------------------------------
//--                                                   --
//--      Base class:  TDynamicDll                     --
//--                                                   --
//-------------------------------------------------------

type
  TDynamicDll = class(TComponent)
  protected
    function IsAPIVersionStored: Boolean;
    function IsDllNameStored: Boolean;
    function IsRegVersionStored: Boolean;
    function GetDllPath : string; virtual;
    function GetDllName: string; virtual;
    procedure SetDllName(const Value: string); virtual;
    procedure SetDllPath(const Value: string); virtual;
  protected
    FDllName            : string;
    FDllPath            : string;
    FAPIVersion         : Integer;
    FRegVersion         : string;
    FLastError          : string;
    FAutoLoad           : Boolean;
    FAutoUnload         : Boolean;
    FFatalMsgDlg        : Boolean;
    FFatalAbort         : Boolean;
    FDLLHandle          : THandle;
    FUseLastKnownVersion: Boolean;
    FOnBeforeLoad       : TNotifyEvent;
    FOnAfterLoad        : TNotifyEvent;
    FOnBeforeUnload     : TNotifyEvent;

    procedure CallMapDll; virtual;
    function  Import(const funcname: string; ExceptionOnFailure: Boolean = True): Pointer;
    function  Import2(funcname: string; args: integer=-1; ExceptionOnFailure: Boolean = True): Pointer;
    procedure Loaded; override;
    procedure BeforeLoad; virtual;
    procedure AfterLoad; virtual;
    procedure BeforeUnload; virtual;
    function  GetQuitMessage : string; virtual;
    procedure DoOpenDll(const aDllName : string); virtual;
    procedure MapDll; virtual; abstract;

  public
    // Constructors & Destructors
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy;                    override;

    // Public methods
    procedure OpenDll(const aDllName : string);
    function  IsHandleValid : Boolean;
    function LoadDll: Boolean; virtual;
    procedure UnloadDll;
    procedure Quit;
    function GetDllFullFileName: String;
    class function CreateInstance(DllPath: String = ''; DllName: String = ''): TDynamicDll;
    class function CreateInstanceAndLoad(DllPath: String = ''; DllName: String = ''): TDynamicDll;

    // Public properties
  published
    property AutoLoad : Boolean read FAutoLoad write FAutoLoad default True;
    property AutoUnload : Boolean read FAutoUnload write FAutoUnload default True;
    property DllName : string read GetDllName write SetDllName stored IsDllNameStored;
    property DllPath : string read GetDllPath write SetDllPath;
    property DllFullFileName : String read GetDllFullFileName;
    property LastError : String read FLastError;
    property APIVersion : Integer read FAPIVersion write FAPIVersion stored IsAPIVersionStored;
    property RegVersion : string read FRegVersion write FRegVersion stored IsRegVersionStored;
    property FatalAbort :  Boolean read FFatalAbort write FFatalAbort default True;
    property FatalMsgDlg : Boolean read FFatalMsgDlg write FFatalMsgDlg default True;
    property UseLastKnownVersion: Boolean read FUseLastKnownVersion write FUseLastKnownVersion default True;
    property OnAfterLoad : TNotifyEvent read FOnAfterLoad write FOnAfterLoad;
    property OnBeforeLoad : TNotifyEvent read FOnBeforeLoad write FOnBeforeLoad;
    property OnBeforeUnload : TNotifyEvent read FOnBeforeUnload write FOnBeforeUnload;
  end;

implementation

(*******************************************************)
(**                                                   **)
(**            class TDynamicDll                      **)
(**                                                   **)
(*******************************************************)

procedure TDynamicDll.DoOpenDll(const aDllName : string);
var
  _DllName: string;
begin
  if not IsHandleValid then
  begin
    if aDllName<>'' then
      _DllName := aDllName
    else
      _DllName := DllName;
    if _DllName='' then
      exit;
    {$IFDEF MSWINDOWS}
    SetDllDirectory(PChar(GetDllPath));
    FDLLHandle := SafeLoadLibrary(
      {$IFDEF FPC}
      PAnsiChar(AnsiString(_DllName))
      {$ELSE}
      _DllName
      {$ENDIF});
    {$ELSE}
    //Linux: need here RTLD_GLOBAL, so Python can do "import ctypes"
    FDLLHandle := THandle(dlopen(PAnsiChar(AnsiString(GetDllPath+_DllName)),
      RTLD_LAZY+RTLD_GLOBAL));
    {$ENDIF}
  end;
end;

function  TDynamicDll.GetDllPath : string;
begin
  Result := FDllPath;
end;

procedure TDynamicDll.SetDllPath(const Value: String);
begin
  if Value='' then
    FDllPath := ''
  else
    FDllPath := IncludeTrailingPathDelimiter(Value);
end;

function TDynamicDll.GetDllFullFileName: String;
begin
  Result := DllPath + DllName;
end;

procedure  TDynamicDll.OpenDll(const aDllName : string);
begin
  UnloadDll;

  BeforeLoad;

  FDLLHandle := 0;

  DoOpenDll(aDllName);

  if not IsHandleValid then begin
    {$IFDEF MSWINDOWS}
    FLastError := Format('Error %d: Could not open Dll "%s"',[GetLastError, DllName]);
    {$ELSE}
    FLastError := Format('Error: Could not open Dll "%s"',[DllName]);
    {$ENDIF}
    if FatalMsgDlg then
      {$IFDEF MSWINDOWS}
      MessageBox( GetActiveWindow, PChar(FLastError), 'Error', MB_TASKMODAL or MB_ICONSTOP );
      {$ELSE}
      WriteLn(ErrOutput, FLastError);
      {$ENDIF}

    if FatalAbort then
      Quit;
  end else
    AfterLoad;
end;

constructor TDynamicDll.Create(AOwner: TComponent);
begin
  inherited;
  FFatalMsgDlg          := True;
  FFatalAbort           := True;
  FAutoLoad             := True;
  FUseLastKnownVersion  := True;
  FDLLHandle            := 0;
end;

destructor TDynamicDll.Destroy;
begin
  if AutoUnload then
    UnloadDll;
  inherited;
end;

function TDynamicDll.Import(const funcname: string; ExceptionOnFailure : Boolean = True): Pointer;
var
  E : EDllImportError;
begin
  Result := GetProcAddress( FDLLHandle, PChar(funcname) );
  if (Result = nil) and ExceptionOnFailure then begin
    {$IFDEF MSWINDOWS}
    E := EDllImportError.CreateFmt('Error %d: could not map symbol "%s"', [GetLastError, funcname]);
    E.ErrorCode := GetLastError;
    {$ELSE}
    E := EDllImportError.CreateFmt('Error: could not map symbol "%s"', [funcname]);
    {$ENDIF}
    E.WrongFunc := funcname;
    raise E;
  end;
end;

function TDynamicDll.Import2(funcname: string; args: integer; ExceptionOnFailure: Boolean): Pointer;
begin
  {$IFDEF WIN32}
  // using STDCall name decoration
  // copy paste the function names from dependency walker to notepad and search for the function name there.
  if args>=0 then
    funcname := '_'+funcname+'@'+IntToStr(args);
  {$ENDIF}
  Result := Import(funcname, ExceptionOnFailure);
end;


procedure TDynamicDll.Loaded;
begin
  inherited;
  if AutoLoad and not (csDesigning in ComponentState) then
    LoadDll;
end;

function  TDynamicDll.IsHandleValid : Boolean;
begin
{$IFDEF MSWINDOWS}
  Result := (FDLLHandle >= 32);
{$ELSE}
  Result := FDLLHandle <> 0;
{$ENDIF}
end;

function TDynamicDll.LoadDll: Boolean;
begin
  OpenDll( DllName );
  Result := IsHandleValid;
end;

procedure TDynamicDll.UnloadDll;
begin
  if IsHandleValid then begin
    BeforeUnload;
    FreeLibrary(FDLLHandle);
    FDLLHandle := 0;
  end;
end;

procedure TDynamicDll.BeforeLoad;
begin
  if Assigned( FOnBeforeLoad ) then
    FOnBeforeLoad( Self );
end;

procedure TDynamicDll.AfterLoad;
begin
  if Assigned( FOnAfterLoad ) then
    FOnAfterLoad( Self );
  CallMapDll;
end;

procedure TDynamicDll.BeforeUnload;
begin
  if Assigned( FOnBeforeUnload ) then
    FOnBeforeUnload( Self );
end;

function  TDynamicDll.GetQuitMessage : string;
begin
  Result := Format( 'Dll %s could not be loaded. We must quit.', [DllName]);
end;

procedure TDynamicDll.Quit;
begin
  if not( csDesigning in ComponentState ) then begin
{$IFDEF MSWINDOWS}
    MessageBox( GetActiveWindow, PChar(GetQuitMessage), 'Error', MB_TASKMODAL or MB_ICONSTOP );
    ExitProcess( 1 );
{$ELSE}
    WriteLn(ErrOutput, GetQuitMessage);
    Halt( 1 );
{$ENDIF}
  end;
end;

function TDynamicDll.IsAPIVersionStored: Boolean;
begin
  Result := not UseLastKnownVersion;
end;

function TDynamicDll.IsDllNameStored: Boolean;
begin
  Result := not UseLastKnownVersion;
end;

function TDynamicDll.IsRegVersionStored: Boolean;
begin
  Result := not UseLastKnownVersion;
end;

procedure TDynamicDll.SetDllName(const Value: string);
begin
  FDllName := Value;
end;

function TDynamicDll.GetDllName: String;
begin
  Result := FDllName;
end;

procedure TDynamicDll.CallMapDll;
begin
  try
    MapDll;
  except
    on E: Exception do begin
      if FatalMsgDlg then
{$IFDEF MSWINDOWS}
        MessageBox( GetActiveWindow, PChar(E.Message), 'Error', MB_TASKMODAL or MB_ICONSTOP );
{$ELSE}
        WriteLn( ErrOutput, E.Message );
{$ENDIF}
      if FatalAbort then Quit;
    end;
  end;
end;

class function TDynamicDll.CreateInstance(DllPath, DllName: String): TDynamicDll;
begin
  Result := Create(nil);
  if DllPath<>'' then
    Result.DllPath := DllPath;
  if DllName<>'' then
    Result.DllName := DllName;
end;

class function TDynamicDll.CreateInstanceAndLoad(DllPath, DllName: String): TDynamicDll;
begin
  Result := CreateInstance(DllPath, DllName);
  Result.LoadDll;
  if not Result.IsHandleValid then
    FreeAndNil(Result);
end;

end.

