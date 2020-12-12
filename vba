'======================================================================
'
' VBA codes provided by;
' dELAb - Deniz Elektronik Lab.
' www.delab.net
'
'======================================================================
'VBA codes created for Win64 system, needs some modification for Win32:
'--
'Remove 'PtrSafe' prefix to run for Win32 systems:
'                  V
Private Declare PtrSafe Function M_GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Global TusGonderFlag As Boolean
Global Say As Long

Private Sub Workbook_Open()

TusGonderFlag = False   'Flag is False status on startup // Başlangıçta eylem bayrağı aktif değil

End Sub

Sub Basla() 'Start

Sil 'Wipe Out
TusGonderFlag = True    'Flag (Boolean) triggered // Başlatma bayrağı tetiklendi
TusGonder               'Go to KeyStroke subroutine // Tuş basımı alt rutinine git

End Sub

Sub Dur() 'Stop

TusGonderFlag = False   'Stop flag initiated // Durdurma bayrağı tetiklendi
Sil ' Wipe Out // Ekranı Sil/Süpür
End

End Sub

Public Sub TusGonder()  'Send KeyStrokes

    TimeNow = Format(Now, "hh:mm:ss")

While TusGonderFlag = True  'Still flag is true then continue cycle
                            'Bayrak True (1) ise çevrime devam et
    
    If TimeNow <> Format(Now, "hh:mm:ss") Then 'Skip to send same time stamp // Aynı saati göndermeyi geç
    
            DoEvents            'watch the rest events // diğer olayları da gözle
            Check_Excel         'Excel is focused app or not?  // Excel odakta mı?
            SendKeys "Time: " & Format(Now, "hh:mm:ss"), True 'Send current time // Saati gönder
            SendKeys "~", True  'Enter char // Enter karakteri
            Sleep 500           'Wait 0.5 second (500 msec) // 0.5 saniye bekle
            TimeNow = Format(Now, "hh:mm:ss")
            Say = Say + 1      ' Cycle Counter // Çevrim sayacı
            
            If Say = 50 Then   ' wait 50 cycle and wipe out // 50 çevrim bekle ve sonrasında sil
                        Sil 'Wipe Out
                        End If
     End If

Wend

End Sub

Sub Sil()  'Wipe Out // Ekranı sil/süpür
    
    Cells.Select             ' Tüm hücreleri seç
    Selection.ClearContents  ' Tüm içeriği sil
    Range("C7").Select       ' C7'ye git
    Say = 0
    
End Sub

Sub Check_Excel()   'Checks whether Excel is focused app or not // Excel odakta mı değil mi kontrol rutini

    'This part provided from www
    Dim nm As String
    Dim XLwnd As Long
    Dim Hwnd As Long
    nm = Application.Caption ' Excel window title // Excel dosya adı
    XLwnd = FindWindow(CLng(0), nm) ' Excel window number // Excel ekran numarası
    Hwnd = M_GetActiveWindow() ' Which windows is active? // Hangi numaralı ekran aktif?
    If Hwnd = XLwnd Then 'Excel is active so do nothing // Excel aktif ise sorun yok
        Exit Sub
    Else
        'Excel is not active window, to avoid sending keystrokes to another app stop macro immediatelly
        'Excel aktif pencere değil, diğer programa tuş basımı göndermemek için makroyu hemen durdur
        'TR:
        MsgBox "Başka bir uygulamaya geçildiğinden makro durduruldu." & Chr(10) & "Devam etmek için tekrar çalıştırmanız gerekir.", vbCritical, "_UYARI_"
        'EN:
        'MsgBox "Macro stopped because of switching to another application." & Chr(10) & "Run it again to continue.", vbCritical, "_WARNING_"

        End
    End If
    
End Sub







