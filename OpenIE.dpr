program OpenIE;

uses
  Winapi.Windows,
  System.SysUtils,
  uOSUtils,
  uAppStrings in 'uAppStrings.pas',
  uIELoader in 'uIELoader.pas';

{$R *.res}

begin
  if not IsWindowsVersionOrGreater(10, 0, 0) then
  begin
    MessageBox(0, PChar(SWin10RequiredMsg), PChar(SErrorCaption), MB_ICONERROR or MB_OK);
    Halt(1);
  end;

  if ParamCount > 0 then
    IE_OpenWithURL(ParamStr(1))
  else
    IE_OpenGoogle;

  Halt(0);
end.
