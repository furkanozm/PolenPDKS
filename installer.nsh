; Özel NSIS script - PDKS İşleme Sistemi

!macro preInit
  ; Kurulum öncesi işlemler
  SetRegView 64
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "DisplayName" "PDKS İşleme Sistemi"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "Publisher" "Güleryüz Group"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "DisplayVersion" "1.0.0"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}" "NoRepair" 1
!macroend

!macro customInstall
  ; Özel kurulum işlemleri
  CreateDirectory "$INSTDIR\resources"
  CreateDirectory "$INSTDIR\resources\app"
!macroend

!macro customUnInstall
  ; Özel kaldırma işlemleri
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTALL_APP_KEY}"
  
  ; Kullanıcı verilerini koru (isteğe bağlı)
  MessageBox MB_YESNO "Kullanıcı ayarları ve verileri silinsin mi?" IDNO skip_user_data
  RMDir /r "$APPDATA\polenpdks"
  skip_user_data:
!macroend

!macro customHeader
  ; Özel header
  !system "echo PDKS İşleme Sistemi Kurulumu"
!macroend
