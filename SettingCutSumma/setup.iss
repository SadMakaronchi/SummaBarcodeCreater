; ==============================
;  SummaBarcodeCreater Addon
; ==============================

#define MyAppName "SummaBarcodeCreater"
#define MyAppVersion "1.0"
#define MyAppPublisher "SadMakaronchi"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={tmp}
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=no

OutputDir=.
OutputBaseFilename=Setup_Addon
WizardStyle=modern

ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"


[Code]

var
  CorelVersions: array of string;
  CorelPaths: array of string;
  SelectedCorelPath: string;
  CorelPage: TInputOptionWizardPage;

{ ---------------------------------
  Поиск CorelDRAW
  --------------------------------- }
function FindCorelDraw(): Boolean;
var
  Versions: array[0..7] of string;
  i: Integer;
  InstallDir: string;
begin
  Result := False;

  Versions[0] := '19.0'; // X7
  Versions[1] := '20.0'; // X8
  Versions[2] := '21.0'; // 2019
  Versions[3] := '22.0'; // 2020
  Versions[4] := '23.0'; // 2021
  Versions[5] := '24.0'; // 2022
  Versions[6] := '25.0'; // 2023
  Versions[7] := '26.0'; // 2024+

  for i := 0 to GetArrayLength(Versions) - 1 do
  begin
    InstallDir := '';

    if not RegQueryStringValue(
         HKLM64,
         'SOFTWARE\Corel\CorelDRAW\' + Versions[i],
         'ProgramsDir',
         InstallDir
       ) then
    begin
      RegQueryStringValue(
        HKLM64,
        'SOFTWARE\Corel\CorelDRAW\' + Versions[i],
        'InstallDir',
        InstallDir
      );
    end;

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

{ ---------------------------------
  Создание страницы выбора версии
  --------------------------------- }
procedure InitializeWizard();
var
  i: Integer;
begin
  CorelPage := CreateInputOptionPage(
    wpWelcome,
    'Версия CorelDRAW',
    'Выбор версии CorelDRAW',
    'Выберите версию CorelDRAW для установки аддона:',
    True,
    False
  );

  for i := 0 to GetArrayLength(CorelVersions) - 1 do
    CorelPage.Add('CorelDRAW ' + CorelVersions[i]);

  { если версия одна — выбрать автоматически }
  if GetArrayLength(CorelVersions) = 1 then
    CorelPage.Values[0] := True;
end;

{ ---------------------------------
  Обработка кнопки "Далее"
  --------------------------------- }
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

    SelectedCorelPath :=
      CorelPaths[CorelPage.SelectedValueIndex];
  end;
end;

{ ---------------------------------
  Проверка перед стартом
  --------------------------------- }
function InitializeSetup(): Boolean;
begin
  if not FindCorelDraw() then
  begin
    MsgBox(
      'CorelDRAW не найден в системе.'#13#10 +
      'Установка невозможна.',
      mbError,
      MB_OK
    );
    Result := False;
    Exit;
  end;

  Result := True;
end;

{ ---------------------------------
  Путь установки аддона
  --------------------------------- }
function GetAddonDir(): string;
begin
  Result := SelectedCorelPath + '\Programs64\Addons';
end;

{ ---------------------------------
  Копирование файлов аддона
  --------------------------------- }
procedure CopyAddonFiles();
var
  ResultCode: Integer;
begin
  ForceDirectories(GetAddonDir());

  Exec(
    'cmd.exe',
    '/C xcopy "' + ExpandConstant('{src}\Files\*') +
    '" "' + GetAddonDir() + '\" /E /I /Y',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  );
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    CopyAddonFiles();
end;
