unit uExt;

interface

uses
  Winapi.Windows, Winapi.KnownFolders, System.SysUtils, System.Variants, ActiveX,
  ComObj, ShellAPI, SHDocVw, ShlObj;

function RtlGetVersion(var RTL_OSVERSIONINFOEXW): LONG; stdcall; external 'ntdll.dll' Name 'RtlGetVersion';
function IsWindowsVersionLower(Major, Minor, Build: DWORD): Boolean;

function Wow64DisableWow64FsRedirection(var OldValue: Pointer): BOOL; stdcall; external 'kernel32.dll';
function Wow64RevertWow64FsRedirection(OldValue: Pointer): BOOL; stdcall; external 'kernel32.dll';

function IE_OpenWithURL(const InputURL: string): Boolean;
procedure IE_OpenGoogle;

function IE_PrepareURL(const InputURL: string): string;
function IE_ActivateWindow(const WB: IWebBrowser2): Boolean;
function IE_CreateNewInstance(out WB: IWebBrowser2): Boolean;
function IE_CreateProcess: Boolean;
function IE_FindInstance(out WB: IWebBrowser2; const TimeMS: DWORD = 3000): Boolean;
function IE_IsValidInstance(const WB: IWebBrowser2): Boolean;

function GetProgramFilesDirX64: string;
function GetProgramFilesDirX86: string;

implementation

function IsWindowsVersionLower(Major, Minor, Build: DWORD): Boolean;
var
  winver: RTL_OSVERSIONINFOEXW;
begin
  FillChar(winver, SizeOf(winver), 0);
  winver.dwOSVersionInfoSize := SizeOf(winver);
  Result := False;
  if RtlGetVersion(winver) = 0 then
  begin
    if winver.dwMajorVersion < Major then
      Exit(True);
    if winver.dwMajorVersion = Major then
    begin
      if winver.dwMinorVersion < Minor then
        Exit(True);
      if winver.dwMinorVersion = Minor then
        Exit(winver.dwBuildNumber < Build);
    end;
  end;
end;

function GetProgramFilesDirX64: string;
var
  Path: PWideChar;
begin
  Result := 'C:\Program Files';
  Path := nil;
  if Succeeded(SHGetKnownFolderPath(FOLDERID_ProgramFilesX64, 0, 0, Path)) then
  try
    Result := Path;
  finally
    CoTaskMemFree(Path);
  end;
end;

function GetProgramFilesDirX86: string;
var
  Path: PWideChar;
begin
  Result := 'C:\Program Files (x86)';
  Path := nil;
  if Succeeded(SHGetKnownFolderPath(FOLDERID_ProgramFilesX86, 0, 0, Path)) then
  try
    Result := Path;
  finally
    CoTaskMemFree(Path);
  end;
end;

function IE_PrepareURL(const InputURL: string): string;
const
  ALLOWED_PREFIXES: array[0..4] of string =
  ('about:', 'mailto:', 'file:', 'javascript:', 'ms-');
var
  ResultURL, LowerCaseURL: string;
  AllowedPrefix: string;
  NeedsHttps: Boolean;
begin
  ResultURL := Trim(InputURL);

  if Pos('://', ResultURL) = 0 then
  begin
    LowerCaseURL := LowerCase(ResultURL);
    NeedsHttps := True;

    for AllowedPrefix in ALLOWED_PREFIXES do
      if Pos(AllowedPrefix, LowerCaseURL) = 1 then
      begin
        NeedsHttps := False;
        Break;
      end;

    if NeedsHttps then
      ResultURL := 'https://' + ResultURL;
  end;
  Result := ResultURL;
end;

function IE_ActivateWindow(const WB: IWebBrowser2): Boolean;
var
  hWndIE: HWND;
begin
  Result := False;
  if WB = nil then Exit;
  hWndIE := HWND(NativeUInt(WB.HWND));
  if hWndIE = 0 then Exit;

  if IsIconic(hWndIE) then
    ShowWindow(hWndIE, SW_RESTORE)
  else
    ShowWindow(hWndIE, SW_SHOW);

  SetForegroundWindow(hWndIE);
  Result := True;
end;

function IE_CreateNewInstance(out WB: IWebBrowser2): Boolean;
var
  hr: HRESULT;
begin
  WB := nil;
  hr := CoCreateInstance(CLASS_InternetExplorer, nil, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, WB);
  Result := Succeeded(hr) and (WB <> nil);
end;

function IE_CreateProcess: Boolean;
var
  OldFsRedirState: Pointer;
  FsRedirDisabled: BOOL;
  ExecInfo: TShellExecuteInfoW;
  IEPath: string;
begin
  FsRedirDisabled := Wow64DisableWow64FsRedirection(OldFsRedirState);
  try
    ZeroMemory(@ExecInfo, SizeOf(ExecInfo));
    ExecInfo.cbSize := SizeOf(ExecInfo);
    ExecInfo.lpVerb := 'open';

    IEPath := IncludeTrailingPathDelimiter(GetProgramFilesDirX64) + 'Internet Explorer\iexplore.exe';
    if not FileExists(IEPath) then
      IEPath := IncludeTrailingPathDelimiter(GetProgramFilesDirX86) + 'Internet Explorer\iexplore.exe';

    if not FileExists(IEPath) then
      Exit(False);

    ExecInfo.lpFile       := PWideChar(IEPath);
    ExecInfo.lpParameters := PWideChar('-embedding -nomerge about:blank');
    ExecInfo.nShow        := SW_SHOWNORMAL;
    Result := ShellExecuteExW(@ExecInfo);
  finally
    if FsRedirDisabled then
      Wow64RevertWow64FsRedirection(OldFsRedirState);
  end;
end;

function IE_FindInstance(out WB: IWebBrowser2; const TimeMS: DWORD = 3000): Boolean;
var
  ShellWindows: IShellWindows;
  StartTick: DWORD;
  i: Integer;
  Disp: IDispatch;
  VIdx: OleVariant;
begin
  Result := False;
  WB := nil;
  try
    ShellWindows := CreateComObject(CLASS_ShellWindows) as IShellWindows;
  except
    Exit(False);
  end;
  StartTick := GetTickCount;
  repeat
    try
      for i := 0 to ShellWindows.Count - 1 do
      begin
        VIdx := i;
        Disp := ShellWindows.Item(VIdx);
        if (Disp <> nil) and Supports(Disp, IWebBrowser2, WB) then
        begin
          try
            if (WB <> nil) and (WB.HWND <> 0) and IE_IsValidInstance(WB) then
              Exit(True);
          except
            WB := nil;
          end;
        end;
      end;
    except
    end;
    Sleep(80);
  until (GetTickCount - StartTick >= TimeMS);
end;

function IE_IsValidInstance(const WB: IWebBrowser2): Boolean;
var
  hWndIE: HWND;
  ClassBuf: array[0..15] of Char;
  ExePath, IE64, IE86: string;
begin
  Result := False;
  if WB = nil then Exit;

  hWndIE := HWND(NativeUInt(WB.HWND));
  if (hWndIE = 0) or
     (GetClassName(hWndIE, ClassBuf, Length(ClassBuf)) = 0) or
     (not SameText(ClassBuf, 'IEFrame')) then Exit;

  ExePath := WB.FullName;
  if not SameText(ExtractFileName(ExePath), 'iexplore.exe') then Exit;

  IE64 := IncludeTrailingPathDelimiter(GetProgramFilesDirX64) + 'Internet Explorer\iexplore.exe';
  IE86 := IncludeTrailingPathDelimiter(GetProgramFilesDirX86) + 'Internet Explorer\iexplore.exe';

  Result := SameText(ExePath, IE64) or SameText(ExePath, IE86);
end;

function IE_OpenWithURL(const InputURL: string): Boolean;
var
  PreparedURL: string;
  IEBrowser: IWebBrowser2;
  ParamEmpty: OleVariant;
begin
  Result := False;
  PreparedURL := IE_PrepareURL(InputURL);
  OleCheck(CoInitializeEx(nil, COINIT_APARTMENTTHREADED));
  try
    if IE_CreateNewInstance(IEBrowser) then
    begin
      IEBrowser.Visible := True;
      ParamEmpty := EmptyParam;
      IEBrowser.Navigate(PreparedURL, ParamEmpty, ParamEmpty, ParamEmpty, ParamEmpty);
      Result := IE_ActivateWindow(IEBrowser);
      Exit;
    end;

    if IE_CreateProcess and IE_FindInstance(IEBrowser, 3000) then
    begin
      IEBrowser.Visible := True;
      ParamEmpty := EmptyParam;
      IEBrowser.Navigate(PreparedURL, ParamEmpty, ParamEmpty, ParamEmpty, ParamEmpty);
      Result := IE_ActivateWindow(IEBrowser);
    end
    else
    begin
      MessageBox(0, 'Internet Explorer 가 존재하지 않거나 실행할 수 없습니다.', 'Error', MB_ICONERROR or MB_OK);
    end;
  finally
    CoUninitialize;
  end;
end;

procedure IE_OpenGoogle;
begin
  IE_OpenWithURL('https://www.google.com');
end;

end.
