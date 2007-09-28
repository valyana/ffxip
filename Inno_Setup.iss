; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!



[Setup]
AppName=FFXI Parser
AppVerName=Version 6.7.1
AppPublisher=Spyle
AppPublisherURL=http://www.frontiernet.net/~Spyle/FFXI/ffxi.html
AppSupportURL=http://www.frontiernet.net/~Spyle/FFXI/ffxi.html
AppUpdatesURL=http://www.frontiernet.net/~Spyle/FFXI/ffxi.html
DefaultDirName={pf}\FFXIP
DefaultGroupName=FFXI Parser
Compression=lzma
SolidCompression=yes
ChangesAssociations=yes

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked


[Registry]
Root: HKCR; Subkey: ".prs"; ValueType: string; ValueName: ""; ValueData: "FFXIParse"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "FFXIParse"; ValueType: string; ValueName: ""; ValueData: "FFXIP File"; Flags: uninsdeletekey
Root: HKCR; Subkey: "FFXIParse\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\FFXI_Parser.EXE,0"
Root: HKCR; Subkey: "FFXIParse\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\FFXI_Parser.EXE"" ""%1"""

[Files]
Source: "c:\vbfiles\scr56en.exe"; DestDir: "{tmp}"; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01
Source: "c:\vbfiles\scripten.exe"; DestDir: "{tmp}"; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02
Source: "c:\projects\strings\FFXI_Parser.exe"; DestDir: "{app}"; Flags: ignoreversion
; begin VB system files
; (Note: Scroll to the right to see the full lines!)
Source: "c:\vbfiles\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "c:\vbfiles\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "c:\vbfiles\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "c:\vbfiles\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "c:\vbfiles\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "c:\vbfiles\comcat.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
; end VB system files
; begin Addtn'l VB Files
Source: "c:\vbfiles\MSINET.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "c:\vbfiles\RICHTX32.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
;Source: "c:\Windows\system32\msvcrt.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver allowunsafefiles
Source: "c:\Windows\system32\comdlg32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver allowunsafefiles
;Source: "c:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\RICHED32.DLL"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver allowunsafefiles
; End Addtn'l VB Files


[INI]
Filename: "{app}\FFXI_Parser.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.frontiernet.net/~Spyle/FFXI/ffxi.html"

[Icons]
Name: "{group}\FFXI Parser"; Filename: "{app}\FFXI_Parser.exe"
Name: "{group}\{cm:ProgramOnTheWeb,FFXI Parser}"; Filename: "{app}\FFXI_Parser.url"
Name: "{group}\{cm:UninstallProgram,FFXI Parser}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\FFXI Parser"; Filename: "{app}\FFXI_Parser.exe"; Tasks: desktopicon

[Run]
; Install Windows 98, Me, and NT 4.0 version
Filename: "{tmp}\scr56en.exe"; Parameters: "/r:n /q:1"; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01
; Install Windows 2000 and XP version
Filename: "{tmp}\scripten.exe"; Parameters: "/r:n /q:1"; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02
Filename: "{app}\FFXI_Parser.exe"; Description: "{cm:LaunchProgram,FFXI Parser}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\FFXI_Parser.url"
