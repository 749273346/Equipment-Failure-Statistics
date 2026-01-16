; Inno Setup Script for Device Defect Statistics Tool
; 建议使用 Inno Setup 6.0 或更高版本编译此脚本

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{A30B1C3D-5E7F-9G1H-I3J5-K7L9M1N3O5P}}
AppName=设备缺陷统计管理系统
AppVersion=2.0
AppPublisher=QC攻关小组
AppCopyright=Copyright (C) 2026 QC攻关小组
DefaultDirName={autopf}\DeviceDefectStats
DefaultGroupName=设备缺陷统计管理系统
AllowNoIcons=yes
; Output location for the installer
OutputDir=.
OutputBaseFilename=DeviceDefectStats_Setup_v2.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Require admin privileges to install to Program Files
PrivilegesRequired=admin
SetupIconFile=..\app_icon.ico
UninstallDisplayIcon={app}\DeviceDefectStats.exe

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; The main executable and dependencies
; NOTE: These paths are relative to the location of this .iss file
Source: "DeviceDefectStats\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Documentation
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\设备缺陷统计管理系统"; Filename: "{app}\DeviceDefectStats.exe"
Name: "{group}\卸载设备缺陷统计管理系统"; Filename: "{uninstallexe}"
Name: "{autodesktop}\设备缺陷统计管理系统"; Filename: "{app}\DeviceDefectStats.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\DeviceDefectStats.exe"; Description: "{cm:LaunchProgram,设备缺陷统计管理系统}"; Flags: nowait postinstall skipifsilent

[Dirs]
; Create the data directory so the user sees where to put files
Name: "{app}\3-设备缺陷问题库及设备缺陷处理记录"

[Code]
// You can add custom Pascal Script code here if needed
