; ==============================
;  SummaBarcodeCreater Addon
; ==============================

#define MyAppName "SummaBarcodeCreater"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "SadMakaronchi"

[Setup]
AppName={#MyAppName}
AppVersion=1.0.0
AppPublisher=SadMakaronchi

DefaultDirName={tmp}
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=no

OutputDir=.
OutputBaseFilename=SummaBarcodeCreaterSetup
WizardStyle=modern

ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
DisableStartupPrompt=False
RestartIfNeededByRun=False

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
Source: "..\source\repos\SettingCutSumma\SettingCutSumma\Addons\SettingCutSumma\CorelDrw.addon"; DestDir: "{code:GetAddonDir}"; Flags: ignoreversion
Source: "..\source\repos\SettingCutSumma\SettingCutSumma\UserUI.xslt"; DestDir: "{code:GetAddonDir}"; Flags: ignoreversion
Source: "\\DESKTOP-S94IK04\MijCtrl\регистрация библиотеки.cmd"; DestDir: "{code:GetAddonDir}"; Flags: ignoreversion
Source: "..\source\repos\SadMakaronchi\SummaBarcodeCreater\SettingCutSumma\Addons\SettingCutSumma\SummaMetki.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\source\repos\SadMakaronchi\SummaBarcodeCreater\SettingCutSumma\Addons\SettingCutSumma\AppUI.xslt"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{code:GetAddonDir}\регистрация библиотеки.cmd"

[Code]

const
  WM_CLOSE = $0010;

var
  CorelVersions: array of string;
  CorelPaths: array of string;
  SelectedCorelPath: string;
  CorelPage: TInputOptionWizardPage;
  CorelWasRunning: Boolean;

function FindWindow(lpClassName, lpWindowName: string): HWND;
  external 'FindWindowW@user32.dll stdcall';

function PostMessage(hWnd: HWND; Msg: UINT; wParam, lParam: Longint): BOOL;
  external 'PostMessageW@user32.dll stdcall';

function IsWindow(hWnd: HWND): BOOL;
  external 'IsWindow@user32.dll stdcall';

procedure Sleep(dwMilliseconds: DWORD);
  external 'Sleep@kernel32.dll stdcall';

{ -------- поиск Corel -------- }

function FindCorelDraw(): Boolean;
var
  Versions: array[0..7] of string;
  i: Integer;
  InstallDir: string;
begin
  Result := False;

  Versions[0] := '19.0';
  Versions[1] := '20.0';
  Versions[2] := '21.0';
  Versions[3] := '22.0';
  Versions[4] := '23.0';
  Versions[5] := '24.0';
  Versions[6] := '25.0';
  Versions[7] := '26.0';

  for i := 0 to High(Versions) do
  begin
    InstallDir := '';
    RegQueryStringValue(
      HKLM64,
      'SOFTWARE\Corel\CorelDRAW\' + Versions[i],
      'ProgramsDir',
      InstallDir
    );

    if InstallDir <> '' then
    begin
      SetArrayLength(CorelVersions, GetArrayLength(CorelVersions) + 1);
      SetArrayLength(CorelPaths, GetArrayLength(CorelPaths) + 1);

      CorelVersions[High(CorelVersions)] := Versions[i];
      CorelPaths[High(CorelPaths)] := InstallDir;

      Result := True;
    end;
  end;
end;

{ -------- папка аддона -------- }

function GetAddonDir(Param: string): string;
begin
  Result := SelectedCorelPath + '\Addons\SummaBarcodeCreater';
  ForceDirectories(Result);
end;

{ -------- UI -------- }

procedure InitializeWizard();
var
  i: Integer;
begin
  CorelPage := CreateInputOptionPage(
    wpWelcome,
    'Версия CorelDRAW',
    'Выбор версии CorelDRAW',
    'Выберите версию CorelDRAW:',
    False,
    False
  );

  for i := 0 to GetArrayLength(CorelVersions) - 1 do
    CorelPage.Add('CorelDRAW ' + CorelVersions[i]);

  if GetArrayLength(CorelVersions) = 1 then
    CorelPage.Values[0] := True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;

  if (CorelPage <> nil) and (CurPageID = CorelPage.ID) then
  begin
    if CorelPage.SelectedValueIndex < 0 then
    begin
      MsgBox('Выберите версию CorelDRAW.', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    SelectedCorelPath := CorelPaths[CorelPage.SelectedValueIndex];
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := FindCorelDraw();
  if not Result then
    MsgBox('CorelDRAW не найден.', mbError, MB_OK);
end;








