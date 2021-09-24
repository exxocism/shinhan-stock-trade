Attribute VB_Name = "Module1"
Option Explicit
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As Any, ByVal lpszWindow As Any) As Long
Public Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Const LOAD_ARRAY_ALL As Integer = 0         'LoadArray 함수에서 사용, 셀의 값을 모든 저장소에 로드
Const LOAD_ARRAY_STOCKINFO As Integer = 1   '단순조회 및 실시간조회에 사용되는 값만 로드 (저장소 : stockinfo)
Const LOAD_ARRAY_FAVORITE As Integer = 2    '관심종목과 관련된 셀들을 로드 (저장소 : favoriteinfo)
Const LOAD_ARRAY_BALANCE As Integer = 3     '현재 보유주식과 제비용단가 내용을 로드 (저장소 : balanceinfo)
'                                               - (본 함수 호출시 단순조회와 관심종목을 같이 불러옴)
Const LOAD_ARRAY_AUTODEAL As Integer = 4    '자동거래에 사용되는 셀들을 로드 (저장소 : autodealinfo)
'                                               - (본 함수 호출시 단순조회 셀도 같이 로드됨)
Const LOAD_ARRAY_BUYSELL As Integer = 5     '구매/판매에 사용되는 셀들을 로드 (저장소 : buysellinfo)
'                                               - (본 함수 호출시 단순조회 셀도 같이 로드됨)
Const LOAD_ARRAY_REALTIME As Integer = 6    '실시간 조회를 할지 말지 확인하는 값을 불러옴 (저장소 : stockinfo)
'                                               - (주의 : stockinfo와는 별도의 셀인 realtime_checkbox에 저장됨)

Const NOT_AUTOSTOP = 0
Const AUTOSTOP_GOOD = 1
Const AUTOSTOP_BAD = -1

Const WM_CHAR = &H102
Const WM_SETTEXT = &HC
Const WM_USER = &H400

Const CRAZYBUY_TIME_BUY = True
Const CRAZYBUY_TIME_SELL = False

Dim HANDLE_PLUGIN As Long

Dim CRAZYBUY_code As String
Dim CRAZYBUY_thistime As Boolean
Dim CRAZYBUY_buyprice As Long
Dim CRAZYBUY_sellprice As Long

Public Sub Realtime_Update()

    Sheet1.Writearray (LOAD_ARRAY_STOCKINFO) 'writing warning
    
    If Sheet1.GetSettings("realtime_enabled") = True Then Application.OnTime Sheets("v").Range("realtime_next_runtime").Value2, "Realtime_Update"

End Sub

Public Sub Timer_StockInfo_Tick()

    If Sheet1.CheckRemaining(LOAD_ARRAY_FAVORITE) = 0 And Sheet1.CheckRemaining(LOAD_ARRAY_STOCKINFO) = 0 Then
        Call Sheet1.Writearray(LOAD_ARRAY_STOCKINFO)
        Call Sheet1.Writearray(LOAD_ARRAY_FAVORITE)
        Sheet1.GetCurrentBalance
        Sheet1.SetTimerCount LOAD_ARRAY_STOCKINFO, 0
        Exit Sub
    End If

    Sheet1.SetTimerCount LOAD_ARRAY_STOCKINFO, Sheet1.CheckTimerCount(LOAD_ARRAY_STOCKINFO) + 1
    
    If Sheet1.CheckTimerCount(LOAD_ARRAY_STOCKINFO) > 30 Then
        Call Sheet1.Writearray(LOAD_ARRAY_STOCKINFO)
        Call Sheet1.Writearray(LOAD_ARRAY_FAVORITE)
        Sheet1.GetCurrentBalance
        Sheet1.SetTimerCount LOAD_ARRAY_STOCKINFO, 0
        Exit Sub
    End If
    
    Application.OnTime Sheets("v").Range("crazybuy_next_ontime").Value2, "Timer_StockInfo_Tick"

End Sub

Public Sub crazybuy_assign(ByRef crazybuy_time As Boolean, ByRef code As String, ByRef buyprice As Long, ByRef sellprice As Long)

    CRAZYBUY_thistime = crazybuy_time
    CRAZYBUY_code = code
    CRAZYBUY_buyprice = buyprice
    CRAZYBUY_sellprice = sellprice
    
End Sub

Public Sub crazybuy_onetime()

    With Sheets("v")
        Select Case CRAZYBUY_thistime
            Case CRAZYBUY_TIME_BUY:          Sheet1.BuyStock CRAZYBUY_code, CRAZYBUY_buyprice, CLng(.Range("crazybuy_amount_pertime").Value2), 1, 1, NOT_AUTOSTOP
            Case CRAZYBUY_TIME_SELL:        Sheet1.SellStock CRAZYBUY_code, CRAZYBUY_sellprice, CLng(.Range("crazybuy_amount_pertime").Value2), 1, 1, NOT_AUTOSTOP
        End Select
        
        'MsgBox "suck"
        CRAZYBUY_thistime = Not CRAZYBUY_thistime
        Sheet1.SetSettings "crazybuy_iteration_time", CLng(.Range("crazybuy_iteration_time").Value2) + 1, True
    End With
    
End Sub

Public Sub Crazybuy_TICK()

    With Sheets("v")
        Select Case CRAZYBUY_thistime
            
            Case CRAZYBUY_TIME_BUY:          Sheet1.BuyStock CRAZYBUY_code, CRAZYBUY_buyprice, CLng(.Range("crazybuy_amount_pertime").Value2), 1, 1, NOT_AUTOSTOP
            Case CRAZYBUY_TIME_SELL:        Sheet1.SellStock CRAZYBUY_code, CRAZYBUY_sellprice, CLng(.Range("crazybuy_amount_pertime").Value2), 1, 1, NOT_AUTOSTOP
            
        End Select
    
    
        CRAZYBUY_thistime = Not CRAZYBUY_thistime
        Sheet1.SetSettings "crazybuy_iteration_time", CLng(.Range("crazybuy_iteration_time").Value2) - 1, True
        
        If .Range("crazybuy_iteration_time").Value2 > 0 And .Range("crazybuy_running").Value2 = True Then
            .Range("crazybuy_next_ontime").Calculate
            Application.OnTime .Range("crazybuy_next_ontime").Value2, "Crazybuy_TICK"
        Else
            .Range("crazybuy_running").Value2 = False
        End If
    End With
    
End Sub

Public Sub VERIFY_NOT_ASSIGNED()
    
    Dim FavoriteReadColumn As Integer
    Dim FavoriteWriteColumn As Integer
    Dim FavoriteStart As Integer
    Dim FavoriteFinish As Integer
    Dim FavoriteNameColumn As Integer
    
    Dim StockReadColumn As Integer
    Dim StockWriteColumn As Integer
    Dim StockReadStart As Integer
    Dim StockReadFinish As Integer
    Dim StockNameColumn As Integer
    
    Dim TradeAmountColumn As Integer
    
    Dim tmp As String
    
    'favorite read
    FavoriteReadColumn = Sheet1.GetColumn("FavoriteStart")
    FavoriteStart = Sheet1.GetRow("FavoriteStart")
    FavoriteFinish = Sheet1.GetRow("FavoriteFinish")
    FavoriteWriteColumn = Sheet1.GetColumn("Favorite_target_column")
    FavoriteNameColumn = Sheet1.GetColumn("FavoriteNameColumn")
    
    'stocklist read
    StockReadColumn = Sheet1.GetColumn("StockreadStart")
    StockReadStart = Sheet1.GetRow("StockreadStart")
    StockReadFinish = Sheet1.GetRow("StockreadFinish")
    StockWriteColumn = Sheet1.GetColumn("Stockread_target_column")
    StockNameColumn = Sheet1.GetColumn("StockNameColumn")
    
    TradeAmountColumn = Sheet1.GetColumn("TradeAmountColumn")
    
    'Load stock first
    Dim row As Integer, col As Integer, stockcode As String
    
    With Sheets("Main")
        col = StockReadColumn
        For row = StockReadStart To StockReadFinish
            stockcode = .Cells(row, col).Value2
            If Len(stockcode) = 6 And IsNumeric(stockcode) Then  'row, column
                If Len(.Cells(row, StockNameColumn).Value2) = 0 Then
                    Sheet1.GetStockName stockcode, row, StockNameColumn
                End If
                If Len(.Cells(row, StockWriteColumn).Value2) = 0 Then
                    Sheet1.GetCurrentPrice stockcode, row, StockWriteColumn
                End If
                If Len(.Cells(row, TradeAmountColumn).Value2) = 0 Then
                    Sheet1.GetTradeAmount stockcode, row, TradeAmountColumn
                End If
            End If
        Next row
        
        'Load Favorite
        col = FavoriteReadColumn
        For row = FavoriteStart To FavoriteFinish
            stockcode = .Cells(row, col).Value2
            If Len(stockcode) = 6 And IsNumeric(stockcode) Then  'row, column
                If Len(.Cells(row, FavoriteNameColumn).Value2) = 0 Then
                    Sheet1.GetStockName stockcode, row, FavoriteNameColumn
                End If
                If Len(.Cells(row, FavoriteWriteColumn).Value2) = 0 Then
                    Sheet1.GetCurrentPrice stockcode, row, FavoriteWriteColumn
                End If
            End If
        Next row
    End With
        
End Sub


Sub GetPluginHandle()
    HANDLE_PLUGIN = FindWindow(CLng(0), "FastVBA Plugin")
End Sub

Public Sub SendFunction(ByRef sFname As String)
    
    If HANDLE_PLUGIN = 0 Then
        GetPluginHandle
        If HANDLE_PLUGIN = 0 Then
            MsgBox "플러그인을 찾을 수 없습니다.", vbOKOnly + vbExclamation, "오류"
            Exit Sub
        End If
    End If
    SendMessage HANDLE_PLUGIN, WM_SETTEXT, CLng(0), sFname
    'PostMessage HANDLE_PLUGIN, WM_SETTEXT, CLng(0), sFname
    
End Sub

Public Sub SF_Force(ByRef sFname As String)
    
    If HANDLE_PLUGIN = 0 Then
        GetPluginHandle
        If HANDLE_PLUGIN = 0 Then
            MsgBox "플러그인을 찾을 수 없습니다.", vbOKOnly + vbExclamation, "오류"
            Exit Sub
        End If
    End If
    SendMessage HANDLE_PLUGIN, WM_SETTEXT, CLng(0), sFname
    'PostMessage HANDLE_PLUGIN, WM_SETTEXT, CLng(0), sFname
    
End Sub

Public Sub QuitApp()
    
    On Error Resume Next
    Sheet1.objShinhan.UnRequestRTRegAll
    Sheet1.objShinhan_balance.UnRequestRTRegAll
    Sheet1.objShinhan_buysell.UnRequestRTRegAll
    Sheet1.objShinhan_RT.UnRequestRTRegAll
    Sheet1.objShinhan_RT2.UnRequestRTRegAll
    Sheet1.objShinhan_RT3.UnRequestRTRegAll
    Sheet1.objShinhan_RT4.UnRequestRTRegAll
    Sheet1.objShinhan_RT5.UnRequestRTRegAll
    Sheet1.objShinhan_RT6.UnRequestRTRegAll
    Sheet1.objShinhan_RT7.UnRequestRTRegAll
    Sheet1.objShinhan_RT8.UnRequestRTRegAll
    Sheet1.objShinhan_RT9.UnRequestRTRegAll
    Sheet1.objShinhan_RT10.UnRequestRTRegAll
    Sheet1.objShinhan_RT11.UnRequestRTRegAll
    Sheet1.objShinhan_RT12.UnRequestRTRegAll
    Sheet1.objShinhan_RT13.UnRequestRTRegAll
    Sheet1.objShinhan_RT14.UnRequestRTRegAll
    Sheet1.objShinhan_RT15.UnRequestRTRegAll
    Sheet1.objShinhan_RT16.UnRequestRTRegAll
    Sheet1.objShinhan_RT17.UnRequestRTRegAll
    Sheet1.objShinhan_RT18.UnRequestRTRegAll
    Sheet1.objShinhan_RT19.UnRequestRTRegAll
    Sheet1.objShinhan_RT20.UnRequestRTRegAll
    
    Sheet1.oShin_PriceAmount.UnRequestRTRegAll
    Sheet1.oShin_PriceOnly.UnRequestRTRegAll
    Sheet1.oShin_name.UnRequestRTRegAll
    Sheet1.oShin_favname.UnRequestRTRegAll


    Sheet1.objShinhan.SelfMemFree True
    Sheet1.objShinhan_balance.SelfMemFree True
    Sheet1.objShinhan_buysell.SelfMemFree True
    Sheet1.objShinhan_RT.SelfMemFree True
    Sheet1.objShinhan_RT2.SelfMemFree True
    Sheet1.objShinhan_RT3.SelfMemFree True
    Sheet1.objShinhan_RT4.SelfMemFree True
    Sheet1.objShinhan_RT5.SelfMemFree True
    Sheet1.objShinhan_RT6.SelfMemFree True
    Sheet1.objShinhan_RT7.SelfMemFree True
    Sheet1.objShinhan_RT8.SelfMemFree True
    Sheet1.objShinhan_RT9.SelfMemFree True
    Sheet1.objShinhan_RT10.SelfMemFree True
    Sheet1.objShinhan_RT11.SelfMemFree True
    Sheet1.objShinhan_RT12.SelfMemFree True
    Sheet1.objShinhan_RT13.SelfMemFree True
    Sheet1.objShinhan_RT14.SelfMemFree True
    Sheet1.objShinhan_RT15.SelfMemFree True
    Sheet1.objShinhan_RT16.SelfMemFree True
    Sheet1.objShinhan_RT17.SelfMemFree True
    Sheet1.objShinhan_RT18.SelfMemFree True
    Sheet1.objShinhan_RT19.SelfMemFree True
    Sheet1.objShinhan_RT20.SelfMemFree True
    
    Sheet1.oShin_PriceAmount.SelfMemFree True
    Sheet1.oShin_PriceOnly.SelfMemFree True
    Sheet1.oShin_name.SelfMemFree True
    Sheet1.oShin_favname.SelfMemFree True
    
    Set Sheet1.objShinhan = Nothing
    Set Sheet1.objShinhan_balance = Nothing
    Set Sheet1.objShinhan_buysell = Nothing
    Set Sheet1.objShinhan_RT = Nothing
    Set Sheet1.objShinhan_RT2 = Nothing
    Set Sheet1.objShinhan_RT3 = Nothing
    Set Sheet1.objShinhan_RT4 = Nothing
    Set Sheet1.objShinhan_RT5 = Nothing
    Set Sheet1.objShinhan_RT6 = Nothing
    Set Sheet1.objShinhan_RT7 = Nothing
    Set Sheet1.objShinhan_RT8 = Nothing
    Set Sheet1.objShinhan_RT9 = Nothing
    Set Sheet1.objShinhan_RT10 = Nothing
    Set Sheet1.objShinhan_RT11 = Nothing
    Set Sheet1.objShinhan_RT12 = Nothing
    Set Sheet1.objShinhan_RT13 = Nothing
    Set Sheet1.objShinhan_RT14 = Nothing
    Set Sheet1.objShinhan_RT15 = Nothing
    Set Sheet1.objShinhan_RT16 = Nothing
    Set Sheet1.objShinhan_RT17 = Nothing
    Set Sheet1.objShinhan_RT18 = Nothing
    Set Sheet1.objShinhan_RT19 = Nothing
    Set Sheet1.objShinhan_RT20 = Nothing
    
    Set Sheet1.oShin_PriceAmount = Nothing
    Set Sheet1.oShin_PriceOnly = Nothing
    Set Sheet1.oShin_name = Nothing
    Set Sheet1.oShin_favname = Nothing
    
    SF_Force ("Quit_APP")
    
End Sub

Sub test()
    MsgBox "stub"
End Sub
