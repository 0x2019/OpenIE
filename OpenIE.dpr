program OpenIE;

uses
  Winapi.Windows, System.SysUtils,
  uExt in 'uExt.pas';

{$O+} {$SetPEFlags IMAGE_FILE_RELOCS_STRIPPED}
{$R *.res}

begin
  if IsWindowsVersionLower(10, 0, 0) then
  begin
    MessageBox(0, '이 프로그램은 Windows 10 이상에서만 실행됩니다.', 'Error', MB_ICONERROR or MB_OK);
    Halt(1);
  end;

  if ParamCount > 0 then
    IE_OpenWithURL(ParamStr(1))
  else
    IE_OpenGoogle;

  Halt(0);
end.
