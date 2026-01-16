; *** Inno Setup version 6.0.0+ Chinese Simplified messages ***
;
[LangOptions]
LanguageName=Chinese Simplified
LanguageID=$0804
LanguageCodePage=65001

[Messages]
; *** Application titles
SetupAppTitle=安装
SetupWindowTitle=安装 - %1
UninstallAppTitle=卸载
UninstallAppFullTitle=%1 卸载

; *** Misc. common
InformationTitle=信息
ConfirmTitle=确认
ErrorTitle=错误

; *** SetupLdr messages
SetupLdrStartupMessage=这将安装 %1。是否继续？
LdrCannotCreateTemp=无法创建临时文件。安装中止
LdrCannotExecTemp=无法执行临时目录中的文件。安装中止

; *** Startup error messages
LastErrorMessage=%1.%n%n错误 %2: %3
SetupFileMissing=安装目录中的文件 %1 丢失。请更正这个问题或获取一个新的程序副本。
SetupFileCorrupt=文件已损坏。请获取一个新的程序副本。
SetupFileCorruptOrWrongVer=文件已损坏，或者与此版本的安装程序不兼容。请更正这个问题或获取一个新的程序副本。
InvalidParameter=传递了无效的命令行参数:%n%n%1
SetupAlreadyRunning=安装程序正在运行。
WindowsVersionNotSupported=该程序不支持您计算机上运行的 Windows 版本。
WindowsServicePackRequired=该程序要求 %1 Service Pack %2 或更高版本。
NotOnThisPlatform=该程序将无法在 %1 上运行。
OnlyOnThisPlatform=该程序必须在 %1 上运行。
OnlyOnTheseArchitectures=该程序只能在专为以下处理器架构设计的 Windows 版本上安装:%n%n%1
WinVersionTooLowError=该程序需要 %1 版本 %2 或更高版本。
WinVersionTooHighError=该程序无法安装在 %1 版本 %2 或更高版本上。
AdminPrivilegesRequired=登录时必须作为管理员安装此程序。
PowerUserPrivilegesRequired=登录时必须作为管理员或有权用户组(Power Users)成员安装此程序。
SetupAppRunningError=安装程序检测到 %1 当前正在运行。%n%n请先关闭它的所有实例，然后单击“确定”继续，或单击“取消”退出。
UninstallAppRunningError=卸载程序检测到 %1 当前正在运行。%n%n请先关闭它的所有实例，然后单击“确定”继续，或单击“取消”退出。

; *** Startup questions
PrivilegesRequiredOverrideTitle=选择安装模式
PrivilegesRequiredOverrideInstruction=选择安装模式
PrivilegesRequiredOverrideText1=%1 可以为所有用户安装(需要管理权限)，或仅为您安装。
PrivilegesRequiredOverrideText2=%1 只能为您安装，或只能为所有用户安装(需要管理权限)。
PrivilegesRequiredOverrideAllUsers=为所有用户安装
PrivilegesRequiredOverrideAllUsersRecommended=为所有用户安装(推荐)
PrivilegesRequiredOverrideCurrentUser=仅为我安装
PrivilegesRequiredOverrideCurrentUserRecommended=仅为我安装(推荐)

; *** Misc. errors
ErrorCreatingDir=安装程序无法创建目录“%1”
ErrorTooManyFilesInDir=无法在目录“%1”中创建文件，因为它包含的文件太多

; *** Setup common messages
ExitSetupTitle=退出安装程序
ExitSetupMessage=安装尚未完成。如果现在退出，程序将不会被安装。%n%n您可以以后再运行安装程序完成安装。%n%n退出安装程序吗？
AboutSetupMenuItem=关于安装程序(&A)...
AboutSetupTitle=关于安装程序
AboutSetupMessage=%1 版本 %2%n%3%n%n%1 主页:%n%4
AboutSetupNote=
TranslatorNote=

; *** Buttons
ButtonBack=< 上一步(&B)
ButtonNext=下一步(&N) >
ButtonInstall=安装(&I)
ButtonOK=确定
ButtonCancel=取消
ButtonYes=是(&Y)
ButtonYesToAll=全是(&A)
ButtonNo=否(&N)
ButtonNoToAll=全否(&O)
ButtonFinish=完成(&F)
ButtonBrowse=浏览(&B)...
ButtonWizardBrowse=浏览(&R)...
ButtonNewFolder=新建文件夹(&M)

; *** "Select Language" dialog messages
SelectLanguageTitle=选择安装语言
SelectLanguageLabel=选择安装过程中使用的语言。

; *** Common wizard text
ClickNext=单击“下一步”继续，或单击“取消”退出安装程序。
BeveledLabel=
BrowseDialogTitle=浏览文件夹
BrowseDialogLabel=在下面的列表中选择一个目录，然后单击“确定”。
NewFolderName=新文件夹

; *** "Welcome" wizard page
WelcomeLabel1=欢迎使用 [name] 安装向导
WelcomeLabel2=将在您的计算机上安装 [name/ver]。%n%n建议您在继续之前关闭所有其他应用程序。

; *** "Password" wizard page
WizardPassword=密码
PasswordLabel1=此安装受密码保护。
PasswordLabel3=请输入密码，然后单击“下一步”继续。密码区分大小写。
PasswordEditLabel=密码(&P):
IncorrectPassword=输入的密码不正确。请重试。

; *** "License Agreement" wizard page
WizardLicense=许可协议
LicenseLabel=继续之前，请阅读以下重要信息。
LicenseLabel3=请阅读以下许可协议。您必须接受此协议的条款，然后才能继续安装。
LicenseAccepted=我接受协议(&A)
LicenseNotAccepted=我不接受协议(&D)

; *** "Information" wizard pages
WizardInfoBefore=信息
InfoBeforeLabel=继续之前，请阅读以下重要信息。
InfoBeforeClickLabel=准备好继续安装时，单击“下一步”。
WizardInfoAfter=信息
InfoAfterLabel=继续之前，请阅读以下重要信息。
InfoAfterClickLabel=准备好继续安装时，单击“下一步”。

; *** "User Information" wizard page
WizardUserInfo=用户信息
UserInfoDesc=请输入您的信息。
UserInfoName=用户名(&U):
UserInfoOrg=组织(&O):
UserInfoSerial=序列号(&S):
UserInfoNameRequired=您必须输入名称。

; *** "Select Destination Location" wizard page
WizardSelectDir=选择目标位置
SelectDirDesc=将 [name] 安装到哪里？
SelectDirLabel3=安装程序将把 [name] 安装到以下文件夹中。
SelectDirBrowseLabel=若要继续，请单击“下一步”。如果您想选择其他文件夹，请单击“浏览”。
DiskSpaceMBLabel=至少需要有 [mb] MB 的可用磁盘空间。
CannotInstallToNetworkDrive=安装程序无法安装到网络驱动器。
CannotInstallToUNCPath=安装程序无法安装到 UNC 路径。
InvalidPath=您必须输入包含驱动器号的完整路径；例如:%n%nC:\APP%n%n或格式为 UNNC 的网络路径:%n%n\\server\share
InvalidDrive=您选择的驱动器或 UNC 共享不存在或无法访问。请选择其他位置。
DiskSpaceWarningTitle=磁盘空间不足
DiskSpaceWarning=安装程序至少需要 %1 KB 的可用磁盘空间才能安装，但选定驱动器上只有 %2 KB 可用。%n%n无论如何都要继续吗？
DirNameTooLong=文件夹名称或路径太长。
InvalidDirName=文件夹名称无效。
BadDirName32=文件夹名称不能包含以下任何字符:%n%n%1
DirExistsTitle=文件夹已存在
DirExists=文件夹:%n%n%1%n%n已经存在。无论如何都要安装到那个文件夹吗？
DirDoesntExistTitle=文件夹不存在
DirDoesntExist=文件夹:%n%n%1%n%n不存在。需要创建吗？

; *** "Select Components" wizard page
WizardSelectComponents=选择组件
SelectComponentsDesc=应该安装哪些组件？
SelectComponentsLabel2=选择要安装的组件；清除不想安装的组件。准备好继续时，单击“下一步”。
FullInstallation=完全安装
CompactInstallation=精简安装
CustomInstallation=自定义安装
NoUninstallWarningTitle=组件存在
NoUninstallWarning=安装程序检测到已安装以下组件:%n%n%1%n%n取消选择这些组件将不会卸载它们。%n%n无论如何都要继续吗？
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=当前选择至少需要 [mb] MB 的磁盘空间。

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=选择附加任务
SelectTasksDesc=应该执行哪些附加任务？
SelectTasksLabel2=选择安装 [name] 时要执行的附加任务，然后单击“下一步”。

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=选择开始菜单文件夹
SelectStartMenuFolderDesc=安装程序应该在哪里放置程序的快捷方式？
SelectStartMenuFolderLabel3=安装程序将在以下“开始”菜单文件夹中创建程序的快捷方式。
SelectStartMenuFolderBrowseLabel=若要继续，请单击“下一步”。如果您想选择其他文件夹，请单击“浏览”。
MustEnterGroupName=您必须输入文件夹名称。
GroupNameTooLong=文件夹名称或路径太长。
InvalidGroupName=文件夹名称无效。
BadGroupName=文件夹名称不能包含以下任何字符:%n%n%1
NoProgramGroupCheck2=不创建“开始”菜单文件夹(&D)

; *** "Ready to Install" wizard page
WizardReady=准备安装
ReadyLabel1=安装程序现在已准备好开始安装 [name]。
ReadyLabel2a=单击“安装”继续此安装，或单击“上一步”查看或更改任何设置。
ReadyLabel2b=单击“安装”继续此安装。
ReadyMemoUserInfo=用户信息:
ReadyMemoDir=目标位置:
ReadyMemoType=安装类型:
ReadyMemoComponents=选定的组件:
ReadyMemoGroup=开始菜单文件夹:
ReadyMemoTasks=附加任务:

; *** "Preparing to Install" wizard page
WizardPreparing=正在准备安装
PreparingDesc=安装程序正在准备安装 [name]。
PreviousInstallNotCompleted=先前的安装/卸载尚未完成。您需要重新启动计算机才能完成该安装。%n%n重新启动计算机后，请再次运行安装程序以完成 [name] 的安装。
CannotContinue=安装程序无法继续。请单击“取消”退出。
ApplicationsFound=以下应用程序正在使用安装程序需要更新的文件。建议您允许安装程序自动关闭这些应用程序。
ApplicationsFound2=以下应用程序正在使用安装程序需要更新的文件。建议您允许安装程序自动关闭这些应用程序。安装完成后，安装程序将尝试重新启动应用程序。
CloseApplications=自动关闭应用程序(&A)
DontCloseApplications=不要关闭应用程序(&D)
ErrorCloseApplications=安装程序无法自动关闭所有应用程序。建议您在继续之前关闭所有使用安装程序需要更新的文件的应用程序。

; *** "Installing" wizard page
WizardInstalling=正在安装
InstallingLabel=安装程序正在安装 [name]，请稍候。

; *** "Setup Completed" wizard page
FinishedHeadingLabel=[name] 安装向导完成
FinishedLabelNoIcons=安装程序已在您的计算机上安装了 [name]。
FinishedLabel=安装程序已在您的计算机上安装了 [name]。可以通过选择已安装的快捷方式来启动应用程序。
ClickFinish=单击“完成”退出安装程序。
FinishedRestartLabel=要完成 [name] 的安装，安装程序必须重新启动您的计算机。您想现在重新启动吗？
FinishedRestartMessage=要完成 [name] 的安装，安装程序必须重新启动您的计算机。%n%n您想现在重新启动吗？
ShowReadmeCheck=是，我想查看自述文件
YesRadio=是，立即重新启动计算机(&Y)
NoRadio=否，我稍后重新启动计算机(&N)

RunEntryExec=运行 %1
RunEntryShellExec=查看 %1

; *** "Setup Needs the Next Disk"
ChangeDiskTitle=安装程序需要下一个磁盘
SelectDiskLabel2=请插入磁盘 %1 并单击“确定”。%n%n如果此磁盘上的文件可以在不同于下面显示的文件夹中找到，请输入正确的路径或单击“浏览”。
PathLabel=路径(&P):
FileNotInDir2=在“%2”中找不到文件“%1”。请插入正确的磁盘或选择其他文件夹。
SelectDirectoryLabel=请输入下一个磁盘的位置。

; *** Installation phase messages
SetupAborted=安装未完成。%n%n请更正问题并再次运行安装程序。
EntryAbortRetryIgnore=单击“重试”重试，单击“忽略”继续，或单击“中止”取消安装。

; *** Installation status messages
StatusClosingApplications=正在关闭应用程序...
StatusCreateDirs=正在创建目录...
StatusExtractFiles=正在解压文件...
StatusCreateIcons=正在创建快捷方式...
StatusCreateIniEntries=正在创建 INI 条目...
StatusRegisterFiles=正在注册文件...
StatusDeleteFiles=正在删除文件...
StatusRunProgram=正在完成安装...
StatusRestartingApplications=正在重新启动应用程序...
StatusRollback=正在回滚更改...

; *** Misc. errors
ErrorInternal2=内部错误: %1
ErrorFunctionFailedNoCode=%1 失败
ErrorFunctionFailed=%1 失败; 代码 %2
ErrorFunctionFailedWithMessage=%1 失败; 代码 %2.%n%3
ErrorExecutingCode=执行代码时出错:%n%1

; *** Registry errors
ErrorRegOpenKey=打开注册表项出错:%n%1\%2
ErrorRegCreateKey=创建注册表项出错:%n%1\%2
ErrorRegWriteKey=写入注册表项出错:%n%1\%2

; *** INI errors
ErrorIniEntry=在文件“%1”中创建 INI 条目出错。

; *** File copying errors
FileAbortRetryIgnore=单击“重试”重试，单击“忽略”跳过此文件(不推荐)，或单击“中止”取消安装。
FileAbortRetryIgnore2=单击“重试”重试，单击“忽略”继续(不推荐)，或单击“中止”取消安装。
SourceIsCorrupted=源文件已损坏
SourceDoesntExist=源文件“%1”不存在
ExistingFileReadOnly2=无法替换现有文件，因为它被标记为只读。
ExistingFileReadOnlyRetry=单击“重试”以删除只读属性并重试，单击“忽略”以跳过此文件，或单击“中止”以取消安装。
ErrorWritingExisting=试图覆盖现有文件失败:
ErrorRenamingTemp=尝试重命名临时目录中的文件失败:
ErrorCopying=复制文件时出错:

; *** Post-installation errors
ErrorInvalidSever=“%1”不是有效的 .EXE 文件。%n%n无法注册关联。
ErrorRegisterServer=无法注册 DLL/OCX: %1
ErrorUnregisterServer=无法取消注册 DLL/OCX: %1
ErrorRegSvr32Failed=RegSvr32 失败，退出代码 %1

; *** Uninstall messages
UninstallStatusLabel=正在从您的计算机中删除 %1，请稍候。
UninstallCompletionLabel=%1 已成功从您的计算机中删除。
UninstallCompletionLabel2=%1 无法完全删除。%n%n您可以手动删除某些元素。
UninstallOpenError=文件“%1”无法打开。无法卸载
UninstallNotFound=文件“%1”不存在。无法卸载
UninstallUnknownEntry=在卸载日志中遇到未知条目 (%1)
ConfirmUninstall=您确定要完全删除 %1 及其所有组件吗？
UninstallOnlyOnWin64=此安装只能在 64 位 Windows 上卸载。

; *** Uninstaller messages
UninstallIntstalledDate=%1 已安装于 %2
UninstallIntstalledSize=%1 大小 %2

[CustomMessages]
CreateDesktopIcon=创建桌面快捷方式
AdditionalIcons=附加图标
LaunchProgram=运行 %1
