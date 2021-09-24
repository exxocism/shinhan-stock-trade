Option Explicit On
Option Strict On
Imports AxGIEXPERTCONTROLLib
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Threading
Imports System.Timers
'Imports GIExpertControl64Lib

Public Class Form1

    '========================================================================================================
    'VB.Net에서 엑셀을 컨트롤하기 위해 선언한 변수, Microsoft.Office.Interop.Excel 소속 (참조추가 필요)
    '========================================================================================================
    Dim objApp As Application
    Dim objBook As Workbook
    Dim objMain As Worksheet
    Dim objv As Worksheet


    '========================================================================================================
    '요청을 돌리고 모든 요청이 수집되었는지 확인하는 타이머 : 확인이 완료되면 엑셀 시트에 붙여넣기를 한다.
    '========================================================================================================
    'AddHandler() Timer_StockInfo.Tick, AddressOf Timer_StockInfo_Tick
    Dim WithEvents Timer_StockInfo As New System.Timers.Timer() ' 현재가 계산용 : 계산종료시 엑셀 반영, 이후 잔고조회 호출
    Dim WithEvents Timer_Realtime As New System.Timers.Timer()  ' 실시간 가격 엑셀 반영용도
    Dim WithEvents Timer_Buysell As New System.Timers.Timer()  ' 주식구매/판매 및 결과 확인용도 : 계산종료시 엑셀 반영
    '---------------------------------------------------------------------------------------------------------
    Public req_remaining(5) As Integer                             ' 수신되지 않은 요청이 얼마나 있는지 확인. 아래 LOAD_ARRAY 관련 환경변수로 구분
    Public timer_tick_count(5) As Integer                          ' 각 타이머가 얼마나 작동되었는지 기록하는 용도

    '========================================================================================================
    '엑셀 / VB 공히 사용하는 환경 변수
    '========================================================================================================
    Const ORDER_BUY As Integer = 2              '구매/판매를 구분하기 위해 사용하는 변수. 실제 구매요청시에 동일한 값이 사용됨
    Const ORDER_SELL As Integer = 1
    Public Const PRICE_MARKET_PRICE = -1               '지정가와 시장가를 구분하기 위해 사용한 음의 변수
    '---------------------------------------------------------------------------------------------------------
    Const NOT_AUTOSTOP = 0                      '자동구매 여부를 판별하기 위한 환경변수, Buysell_Backend 함수에 활용됨
    Const AUTOSTOP_GOOD = 1
    Const AUTOSTOP_BAD = -1

    '========================================================================================================
    ' VB.Net에서만 사용하는 환경 및 전역변수
    '========================================================================================================
    Const START_OR_STOP_ALL As Integer = -1     'SendMessage를 통한 Excel -> VB.net 호출시 실시간조회 시작/종료 쓰이는 환경변수
    '---------------------------------------------------------------------------------------------------------
    Dim BYEBYE As Boolean                       '프로그램 종료시 호출되는 변수. 에러처리를 하지 않음
    Public AUTODEAL_ENABLED_MASTERKEY As Boolean   '자동거래를 사용할지 말지에 대한 환경변수
    Public Account_number As Integer               'oShin_balance_ReceiveData에서 직접 Excel호출이 되지 않아 별도의 쓰레드로 값을 전달하기 위한 변수
    Public AccountNumString As String               'oShin_balance_ReceiveData에서 직접 Excel호출이 되지 않아 별도의 쓰레드로 값을 전달하기 위한 변수
    Public Account_Balance As Long                 '각각 계좌번호 위치, 잔고를 저장해둠
    '---------------------------------------------------------------------------------------------------------
    Dim EXCEL_REALTIME_CONVERSION_INDEX As Integer  '0~19의 인덱스를 6~35의 값으로 변환해주는 마법의 숫자 시발
    Dim REALTIME_STORAGE() As CRT_And_Autodeal  '실시간 조회를 하기 위한 각 클래스의 저장소 (load_settings에서 실제 엑셀에서 사용하는 만큼 만들어서 호출함)
    '                                            (클래스와 구조체 사용으로 유연하고 적은 코드로 구현이 가능하였음)


    '========================================================================================================
    '엑셀에서 시트를 변경하더라도 각 시트의 절대위치를 잡기 위해 선언한 변수
    '   (VB.net 에서 직접 계산할 수는 없고, Load_settings 간 불러온 뒤 오프셋 확인용도로 사용)
    '   (v 시트에 위치한 값을 불러오며, 행 / 열 추가시 변경해야할 수도 있음)
    '========================================================================================================
    Dim StockReadColumn As Integer      '단순 1회성 조회용 종목번호 열 위치
    Dim StockWriteColumn As Integer     '단순조회 현재가 열 위치
    Dim StockNameColumn As Integer      '단순조회 종목이름 열 위치
    Public StockReadStart As Integer       '단순조회 행 시작위치
    Dim StockReadFinish As Integer      '단순조회 행 종료위치
    '                                       - (단순조회 끝이 아닌 실시간을 포함한 종목 끝임. 행 추가시 반드시 확인 필요!)
    '                                       - (단순조회행의 끝을 보려면 RealtimeStart를 참고해야 함)
    '---------------------------------------------------------------------------------------------------------
    Dim realtime_checkbox_column As Integer '실시간 조회 사용여부 열 위치 (주의 : realtimeinfo에 쓰이지 않음. 관련정보는 realtime_checkbox에 저장, offset산출시 주의)
    '---------------------------------------------------------------------------------------------------------
    Dim RealtimeReadColumn As Integer   '실시간 조회 종목번호 열 위치 (일반조회와 혼용해서 사용되었으므로 코드 변경시 각별한 주의를 요함)
    Dim RealtimeWritecolumn As Integer  '실시간 조회 현재가 열 위치 (일반조회시와 혼용해서 사용되었으므로 코드 변경시 각별한 주의를 요함)
    Dim RealtimeStart As Integer        '실시간 조회 행 시작위치
    Dim RealtimeFinish As Integer       '실시간 조회 행 종료위치 (행 추가시 반드시 확인 필요!)
    '---------------------------------------------------------------------------------------------------------
    Dim TradeAmountColumn As Integer    '실시간 / 단순조회에 같이 쓰이는 거래량 열 위치
    '---------------------------------------------------------------------------------------------------------
    Dim FavoriteReadColumn As Integer   '관심종목 종목번호 열 위치
    Dim FavoriteNameColumn As Integer   '관심종목 종목이름 열 위치
    Dim FavoriteWriteColumn As Integer  '관심종목 현재가 열 위치
    Dim FavoriteStart As Integer        '관심종목 행 시작위치
    Dim FavoriteFinish As Integer       '관심종목 행 종료위치
    '---------------------------------------------------------------------------------------------------------
    Dim Goodprice_deal_Favorite_Column As Integer   '잔고조회 / 제비용단가 : 관심종목란에 사용되는 열 위치
    Dim Goodprice_deal_AUTO_Column As Integer       '잔고조회 / 제비용단가 : 자동거래시에 사용되는 열 위치
    '---------------------------------------------------------------------------------------------------------
    Dim Autodeal_CheckBox_Column As Integer         '자동거래 사용여부 열 위치
    Dim Autodeal_badsell_Price_Col As Integer       '자동거래 손절가격 열 위치
    Dim Autodeal_goodsell_Price_Col As Integer      '자동거래 익절가격 열 위치
    Dim Autodeal_goodpercent_Col As Integer         '자동거래 익절율 계산용 열 위치 (제비용단가 기준)
    Dim Autodeal_badpercent_Col As Integer          '자동거래 손절율 계산용 열 위치 (제비용단가 기준)
    Dim Autodeal_ownstock_col As Integer            '잔고조회 / 자동거래 보유주식 계산용 열 위치
    Dim Autodeal_status_col As Integer              '자동거래 결과 기록용 열 위치
    '---------------------------------------------------------------------------------------------------------
    Dim Order_Ckbox_Column As Integer               '구매/판매 사용여부 열 위치
    Dim Order_price_Column As Integer               '구매/판매가 열 위치 (시장가, ㅅ, 시) 입력가능
    Dim Order_amount_Column As Integer              '구매/판매 주식수량 열 위치 (금액만 입력시 자동 계산됨)
    Dim Order_money_Column As Integer               '구매/판매를 금액으로 할 때 입력하는 열 위치
    Dim Order_result_Column As Integer              '구매/판매 결과를 알려주는 열 위치 (주문송신, 주문중, 주문완료)
    '---------------------------------------------------------------------------------------------------------
    Dim Doctrine_Number_Column As Integer            ' 독트린 번호를 알려주는 열 위치
    Dim Doctrine_Arg1_Column As Integer              ' 독트린 파라메터 1번 위치
    Dim Doctrine_Arg2_Column As Integer              ' 독트린 파라메터 2번 위치


    '========================================================================================================
    ' 엑셀에서 값을 불러와서 임시로 저장해두는 배열과 관련된 변수
    '   (직접 Value2등으로 액세스 할 경우에 성능이 매우 떨어지기 때문에 별도의 저장소를 생성하였음)
    '   (동일한 최적화를 VBA에 적용)
    '   (LoadArray로 불러옴, WriteArray로 쓰기 / 기타 내용은 직접 엑세스)
    '========================================================================================================
    Public Const LOAD_ARRAY_ALL As Integer = 0         'LoadArray 함수에서 사용, 셀의 값을 모든 저장소에 로드
    Public Const LOAD_ARRAY_STOCKINFO As Integer = 1   '단순조회 및 실시간조회에 사용되는 값만 로드 (저장소 : stockinfo)
    Public Const LOAD_ARRAY_FAVORITE As Integer = 2    '관심종목과 관련된 셀들을 로드 (저장소 : favoriteinfo)
    Public Const LOAD_ARRAY_BALANCE As Integer = 3     '현재 보유주식과 제비용단가 내용을 로드 (저장소 : autodealinfo)
    '                                               - (본 함수 호출시 단순조회와 관심종목을 같이 불러옴)
    Public Const LOAD_ARRAY_AUTODEAL As Integer = 4    '자동거래에 사용되는 셀들을 로드 (저장소 : autodealinfo)
    '                                               - (본 함수 호출시 단순조회 셀도 같이 로드됨)
    Public Const LOAD_ARRAY_BUYSELL As Integer = 5     '구매/판매에 사용되는 셀들을 로드 (저장소 : buysellinfo)
    '                                               - (본 함수 호출시 단순조회 셀도 같이 로드됨)
    Public Const LOAD_ARRAY_REALTIME As Integer = 6    '실시간 조회를 할지 말지 확인하는 값을 불러옴 (저장소 : stockinfo)
    '                                               - (주의 : stockinfo와는 별도의 셀인 realtime_checkbox에 저장됨)
    Public Const LOAD_ARRAY_DOCTRINES As Integer = 7   '독트린 관련 내용을 불러옴 (저장소 : doctrineinfo)
    '---------------------------------------------------------------------------------------------------------
    Public stockinfo(,) As Object                  '단순조회 및 실시간조회에 사용
    Public favoriteinfo(,) As Object               '관심종목과 관련된 셀들을 로드
    'Dim balanceinfo(,) As Object                '현재 보유주식과 제비용단가 내용을 로드
    Public autodealinfo(,) As Object               '자동거래에 사용되는 셀들을 로드
    Public buysellinfo(,) As Object                '구매/판매에 사용되는 셀들을 로드
    Public realtime_checkbox(,) As Object          '실시간 조회를 할지 말지 확인하는 값 (주의 : 전체 시작시에만 잠시 활용됨, 상태 확인용으로 사용금지)
    '---------------------------------------------------------------------------------------------------------
    Public doctrineinfo(,) As Object               '독트린 관련 내용 저장소


    '========================================================================================================
    ' 임시로 저장해두는 배열과 실제 엑셀 변수 위치와의 괴리를 해결하기 위해 작성한 별도의 변수
    '   (Load_settings에서 해당 값을 계산해줌, 실제 VB 계산시 해당 위치를 사용)
    '   (행의 시작과 종료는 1 to 배열의 UpperBound를 사용하여 계산)
    '========================================================================================================
    Public isReadCol As Integer                    '단순 1회성 조회용 종목번호 열 위치
    Public isWriteCol As Integer                   '단순조회 현재가 열 위치
    Public isNameCol As Integer                    '단순조회 종목이름 열 위치
    Public isAmountCol As Integer                  '실시간/단순조회 거래량 열 위치
    '---------------------------------------------------------------------------------------------------------
    Dim iFReadCol As Integer                    '관심종목 종목번호 열 위치
    Dim iFNameCol As Integer                    '관심종목 종목이름 열 위치
    Dim iFWriteCol As Integer                   '관심종목 현재가 열 위치
    Dim iFGoodPriceCol As Integer               '잔고조회 / 제비용단가 : 관심종목란에 사용되는 열 위치
    '---------------------------------------------------------------------------------------------------------
    Public iBalGoodPriceCol As Integer             '잔고조회 / 제비용단가 : 자동거래시에 사용되는 열 위치
    Dim iBalOwnAmountCol As Integer             '잔고조회 / 자동거래 보유주식 계산용 열 위치
    '---------------------------------------------------------------------------------------------------------
    Dim iOrderBoxCol As Integer                 '구매/판매 사용여부 열 위치
    Dim iOrderPriceCol As Integer               '구매/판매가 열 위치 (시장가, ㅅ, 시) 입력가능
    Dim iOrderAmountCol As Integer              '구매/판매 주식수량 열 위치 (금액만 입력시 자동 계산됨)
    Dim iOrderMoneyCol As Integer               '구매/판매를 금액으로 할 때 입력하는 열 위치
    Dim iOrderResultCol As Integer              '구매/판매 결과를 알려주는 열 위치 (주문송신, 주문중, 주문완료)
    '---------------------------------------------------------------------------------------------------------
    Dim iDocNumCol As Integer                   ' 독트린 번호를 알려주는 열 위치
    Dim iDocArg1Col As Integer                  ' 독트린 파라메터 1번 위치
    Dim iDocArg2Col As Integer                  ' 독트린 파라메터 2번 위치
    '---------------------------------------------------------------------------------------------------------
    Public iAutoBoxCol As Integer                  '자동거래 사용여부 열 위치
    Public iAutoBadsellCol As Integer              '자동거래 손절가격 열 위치
    Public iAutoGoodsellCol As Integer             '자동거래 익절가격 열 위치
    Public iAutoGoodPerCol As Integer              '자동거래 익절율 계산용 열 위치 (제비용단가 기준)
    Public iAutoBadPerCol As Integer               '자동거래 손절율 계산용 열 위치 (제비용단가 기준)
    Public iAutoOwnCol As Integer                  '잔고조회 / 자동거래 보유주식 계산용 열 위치
    Public iAutoStatusCol As Integer               '자동거래 결과 기록용 열 위치


    '========================================================
    '   어플 시작시 로드되는 함수 (Form1_Load)
    '      * 인디 및 엑셀을 시작하고, 설정을 불러옴   
    '========================================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' 신한아이 인디 시작
        oShin_PriceAmount.StartIndi("redacted", "redacted", "redacted", "C:\SHINHAN-i\indi\giexpertstarter.exe")

        ' 엑셀 불러오기 및 시트 지정
        objApp = New Application()
        objBook = objApp.Workbooks.Open("redacted")
        objApp.Visible = True
        objMain = CType(objBook.Sheets("Main"), Worksheet)
        objv = CType(objBook.Sheets("v"), Worksheet)

        '엑셀 파일의 각종 설정 로드
        Load_settings()

        '데이터그리드 성능 향상코드 적용
        Dim dgvDouble As DataGridViewDoubleBuffer = New DataGridViewDoubleBuffer(DataGridView1)
        dgvDouble.EnableDoubleBuffered()

    End Sub


    '========================================================
    '   어플 종료시 로드되는 함수 (Form1_Closing)
    '      * 각종 컨트롤을 제거하고, 엑셀을 종료함
    '========================================================
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        BYEBYE = True
        '선언했던 모든 OCX를 제거함
        Dim all As CRT_And_Autodeal
        For Each all In REALTIME_STORAGE
            If all.Enabled = True Then all.RT_Stop()
            all.Dispose()
        Next all
        oShin_favname.SelfMemFree(True)
        oShin_name.SelfMemFree(True)
        oShin_Priceonly.SelfMemFree(True)
        oShin_PriceAmount.SelfMemFree(True)
        oShin_favname.Dispose()
        oShin_name.Dispose()
        oShin_Priceonly.Dispose()
        oShin_PriceAmount.Dispose()

        '엑셀의 정상종료를 시도함 : 남아있는경우가 더 많음
        Try
            Dim objBooks As Workbooks = objApp.Workbooks
            objApp.Run("QuitApp")
            objBook.Save()
            objBook.Close(False, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
            objBooks.Close()
            objApp.Quit()
        Catch ex As Exception
            If BYEBYE = False Then MsgBox("byebye error : " & ex.Message)
        End Try

    End Sub


    '========================================================
    '   쓰레기 버튼들
    '      * 아직 그냥 놔둠
    '========================================================
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click




    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        objApp.Run("Sheet1.btnRequest_Click")

    End Sub


    '========================================================
    '   주식정보가 수집되었는지 확인하기 위한 타이머 (Timer_StockInfo_Tick)
    '      * 3초 이상 경과하거나, 모든 정보가 수집된 경우 Writearray 사용하여 엑셀에 값 기록
    '      * 값을 기록한 뒤에 잔고조회 명령을 불러옴
    '      * 타이머 종료시 쓰레드가 종료될 것 우려하여 별도 쓰레드로 불러옴
    '========================================================
    Public Sub Timer_StockInfo_Tick(sender As Object, e As EventArgs) Handles Timer_StockInfo.Elapsed

        If req_remaining(LOAD_ARRAY_FAVORITE) = 0 And req_remaining(LOAD_ARRAY_STOCKINFO) = 0 Then
            Writearray(LOAD_ARRAY_STOCKINFO)
            Writearray(LOAD_ARRAY_FAVORITE)
            Dim thread As New Thread(AddressOf GetBalanceInfo)
            thread.Start()
            Handle_Error("[DEBUG] Stock gathering finished. writing")
            timer_tick_count(LOAD_ARRAY_STOCKINFO) = 0
            Timer_StockInfo.Enabled = False
            Timer_StockInfo.Stop()
            Exit Sub
        End If

        timer_tick_count(LOAD_ARRAY_STOCKINFO) = timer_tick_count(LOAD_ARRAY_STOCKINFO) + 1
        If timer_tick_count(LOAD_ARRAY_STOCKINFO) > 30 Then
            Writearray(LOAD_ARRAY_STOCKINFO)
            Writearray(LOAD_ARRAY_FAVORITE)
            Dim thread As New Thread(AddressOf GetBalanceInfo)
            thread.Start()
            Handle_Error("[WARNING] Failed to get all the information !!! stockinfo")
            timer_tick_count(LOAD_ARRAY_STOCKINFO) = 0
            Timer_StockInfo.Enabled = False
            Timer_StockInfo.Stop()
            Exit Sub
        End If

    End Sub


    '========================================================
    '   실시간 주식정보를 엑셀에 반영하기 위한 타이머 (Timer_Realtime_Elapsed)
    '      * 엑셀에 직접 모든 내용을 기록하는 경우 엑셀이 버티지 못함
    '      * 따라서 일정 주기로 엑셀에는 값만 알려주기 위한 용도로 활용
    '      * 실제 주식확인은 어플 내에서 자동거래 클래스를 통해 확인함
    '========================================================
    Private Sub Timer_Realtime_Elapsed(sender As Object, e As ElapsedEventArgs) Handles Timer_Realtime.Elapsed
        Writearray(LOAD_ARRAY_STOCKINFO) 'writing warning
    End Sub


    '========================================================
    '  구매/판매 관련 주문번호를 따서 반영하기 위한 타이머 (Timer_Realtime_Elapsed)
    '      * 주문 체결결과를 확인하기 전에 셀이 과도하게 업데이트 되는 것을 방지하기 위함
    '      * 주문 3초 뒤 호출, 만약 주문요청이 남아있는 셀이 있는 경우 최신화
    '========================================================
    Private Sub Timer_Buysell_Elapsed(sender As Object, e As ElapsedEventArgs) Handles Timer_Buysell.Elapsed

        If req_remaining(LOAD_ARRAY_BUYSELL) > 0 Then Writearray(LOAD_ARRAY_BUYSELL)
        'timer_tick_count(LOAD_ARRAY_BUYSELL) = 0
        Timer_Buysell.Enabled = False
        Timer_Buysell.Stop()

    End Sub


    '========================================================
    '   SendMessage를 통한 프로세스 간 소통을 잡아내기 위한 intercept함수  (WndProc)
    '      * 64-32비트 포인터 문제로 인해 데이터를 직접 WM_USER 등으로 전달이 제한됨
    '      * 따라서 창 이름을 변경한 뒤 해당 내용을 체크하여 전달하는 식으로 작성
    '      * 미인식을 최대한 막기 위해 엑셀 VBA에 보완코드를 작성 (핸들 저장, 없는경우 한번 더 찾기)
    '      * 추가 함수 호출이 제한되기 때문에 별도 쓰레드를 만들어 값을 전달
    '========================================================
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        MyBase.WndProc(m)
        Select Case m.Msg
            Case &HC
                'Label1.Text = "Sendmessage called. text is : " & Chr(10) & m.ToString() & Chr(10) & "yeah" 'm.GetLParam().ToString
                If Me.Text <> "FastVBA Plugin" Then

                    '디버그 메시지 작성
                    Label1.Text = "[DEBUG] Function Called. Name : " & Me.Text & vbCrLf & Label1.Text
                    Label2.Text = Me.Text
                    Me.Text = "FastVBA Plugin"

                    '함수 이름을 확인하여 각 팡션으로 전달
                    Select Case Label2.Text
                        Case "AutoDeal_Start"
                            Dim thread As New Thread(AddressOf Autodeal_handler)
                            thread.Start(True)
                        Case "AutoDeal_STOP"
                            Dim thread As New Thread(AddressOf Autodeal_handler)
                            thread.Start(False)
                        Case "BuyStock"
                            Dim thread As New Thread(AddressOf Buysell_ALL)
                            thread.Start(CStr(ORDER_BUY))
                        Case "SellStock"
                            Dim thread As New Thread(AddressOf Buysell_ALL)
                            thread.Start(CStr(ORDER_SELL))
                        Case "GetStockInfo"
                            Dim thread As New Thread(AddressOf GetStockInfo)
                            thread.Start()
                        Case "GetBalanceInfo"
                            Dim thread As New Thread(AddressOf GetBalanceInfo)
                            thread.Start()
                        Case "Realtime_StartAll"
                            Dim thread As New Thread(AddressOf Realtime_Start_Handler)
                            thread.Start(CStr(START_OR_STOP_ALL))
                        Case "Realtime_StopAll"
                            Dim thread As New Thread(AddressOf Realtime_Stop_Handler)
                            thread.Start(CStr(START_OR_STOP_ALL))
                        Case "Quit_APP"
                            Dim thread As New Thread(AddressOf Unload_handler)
                            thread.Start()
                        Case Else
                            Dim str() As String
                            If Strings.Left(Label2.Text, 8) = "Realtime" Then
                                str = Split(Label2.Text, "_")
                                Select Case str(1)
                                    Case "True"
                                        Dim thread As New Thread(AddressOf Realtime_Start_Handler)
                                        thread.Start(str(2))
                                    Case "False"
                                        Dim thread As New Thread(AddressOf Realtime_Stop_Handler)
                                        'Dim Parameters = New Object() {}
                                        thread.Start(str(2))
                                End Select
                            End If
                            If Strings.Left(Label2.Text, 8) = "AutoRef_" Then
                                str = Split(Label2.Text, "_")
                                Dim thread As New Thread(AddressOf Autodeal_Refresh)
                                thread.Start(CInt(str(1)))
                            End If
                    End Select
                End If

        End Select

    End Sub

    '========================================================
    '   엑셀 종료시 호출되는 함수  (UnloadMe)
    '      * 폼 제거 및 관련활동 수행
    '========================================================
    Public Sub Unload_handler()
        Try
            BYEBYE = True
            Invoke(New UnloadMe(AddressOf UnloadShit))
        Catch ex As Exception

        End Try
    End Sub

    Delegate Sub UnloadMe()
    Private Sub UnloadShit()
        Me.Close()
        Me.Dispose()
        End
    End Sub

    '========================================================
    '   자동거래 관련 마스터키 활성화  (Autodeal_Handler)
    '      * 관련정보 불러와 저장, 실시간 클래스에서 활용 예정
    '      * 외부 클래스에서 마스터키 상태 확인은 IsAutoDealEnabled 참조
    '========================================================
    Public Sub Autodeal_Handler(params As Object)

        AUTODEAL_ENABLED_MASTERKEY = CBool(params)
        Loadarray(LOAD_ARRAY_AUTODEAL)
        Invoke(New WriteGrid(AddressOf UpdateGrid), New Object() {-1})
        ' if enabled => loadarray and (doctrines)
        Select Case AUTODEAL_ENABLED_MASTERKEY
            Case True
                AccountNumString = CStr(objv.Range("Account_number_" & CStr(objv.Range("Selected_Account").Value2)).Value2)
            Case False

        End Select

    End Sub
    '========================================================
    '   자동거래 최신화  (Autodeal_Refresh)
    '      * 어레이를 불러와서 최신 정보를 담아줌
    '========================================================
    Public Sub Autodeal_Refresh(params As Object)
        AccountNumString = CStr(objv.Range("Account_number_" & CStr(objv.Range("Selected_Account").Value2)).Value2)
        Loadarray(LOAD_ARRAY_AUTODEAL)
        Invoke(New WriteGrid(AddressOf UpdateGrid), New Object() {CInt(params) + EXCEL_REALTIME_CONVERSION_INDEX})
    End Sub

    '========================================================
    '   엑셀에서 불려온 전체구매, 판매를 담당  (Buysell_ALL)
    '      * params를 통하여 구매 / 판매를 구분
    '========================================================
    Public Sub Buysell_ALL(params As Object)

        Dim buyorsell As Integer = CInt(params)
        Dim price As String
        loadarray(LOAD_ARRAY_BUYSELL)
        req_remaining(LOAD_ARRAY_BUYSELL) = 0

        For r As Integer = 1 To buysellinfo.GetUpperBound(0)

            If CBool(buysellinfo(r, iOrderBoxCol)) = False Then Continue For
            price = CStr(buysellinfo(r, iOrderPriceCol))
            If price = vbNullString Then
                Handle_Error("[ERROR] 매수/매도 가격을 입력하십시오. 행 : " & r + StockReadStart - 1)
                'MsgBox("구매가격을 입력하십시오. ", CType(vbOKOnly + vbCritical, MsgBoxStyle), "오류")
                Continue For
            End If

            If CStr(buysellinfo(r, iOrderMoneyCol)) = vbNullString And CStr(buysellinfo(r, iOrderAmountCol)) = vbNullString Then
                Handle_Error("[ERROR] 매수/매도시 금액/주식수 둘중에 하나는 넣어야 합니다. 행 : " & r + StockReadStart - 1)
                'MsgBox("매수/매도시 금액/주식수 둘중에 하나는 넣어야지 ㅄ", CType(vbOKOnly + vbCritical, MsgBoxStyle), "오류")
                Continue For
            End If

            If price = "시장가" Or price = "ㅅ" Or price = "시" Then price = CStr(PRICE_MARKET_PRICE)
            If IsNumeric(price) = False Or (CStr(buysellinfo(r, iOrderAmountCol)) <> vbNullString And IsNumeric(CStr(buysellinfo(r, iOrderAmountCol))) = False) Or (CStr(buysellinfo(r, iOrderMoneyCol)) <> vbNullString and IsNumeric(CStr(buysellinfo(r, iOrderMoneyCol))) = False) Then
                Handle_Error("[ERROR] 매수/매도 행 : " & r + StockReadStart - 1)
                'MsgBox("값을 좆같이 넣어 ㅄ", CType(vbOKOnly + vbCritical, MsgBoxStyle), "오류")
                Continue For
            End If

            If CLng(price) = 0 Or CLng(price) < PRICE_MARKET_PRICE Or CLng(buysellinfo(r, iOrderAmountCol)) < 0 Or CLng(buysellinfo(r, iOrderMoneyCol)) < 0 Then
                Handle_Error("[ERROR] 매수/매도 값을 좆같이 넣어 ㅄ. 행 : " & r + StockReadStart - 1)
                'MsgBox("좆되고 싶냐? ㅄ", CType(vbOKOnly + vbCritical, MsgBoxStyle), "오류")
                Continue For
            End If

            If CLng(price) = PRICE_MARKET_PRICE And CStr(stockinfo(r, isWriteCol)) = vbNullString Then
                Handle_Error("[ERROR] 매수/매도 시장가의 경우 현재가부터 조회해야 합니다. 행 : " & r + StockReadStart - 1)
                'MsgBox("시장가면 먼저 현재가부터 조회 ㅄ", CType(vbOKOnly + vbCritical, MsgBoxStyle), "오류")
                Continue For
                Exit Sub
            End If



            '시장가 아닌경우
            If CLng(price) <> PRICE_MARKET_PRICE Then
                '주만 넣은 경우 금액을 계산
                If CLng(buysellinfo(r, iOrderMoneyCol)) = 0 And CLng(buysellinfo(r, iOrderAmountCol)) > 0 Then
                    buysellinfo(r, iOrderMoneyCol) = CLng(buysellinfo(r, iOrderAmountCol)) * CLng(price)
                    '금액만 넣은 경우 주를 계산
                ElseIf CLng(buysellinfo(r, iOrderMoneyCol)) > 0 And CLng(buysellinfo(r, iOrderAmountCol)) = 0 Then
                    buysellinfo(r, iOrderAmountCol) = CLng(Math.Truncate(CType(buysellinfo(r, iOrderMoneyCol), Decimal) / CType(price, Decimal)))
                End If
            Else
                '시장가인데 현재가가 쓰여 있고 금액만 넣은 경우 금액 / 현재가를 나눈 주를집어넣
                If CLng(stockinfo(r, isWriteCol)) > 0 And CLng(buysellinfo(r, iOrderMoneyCol)) > 0 And CLng(buysellinfo(r, iOrderAmountCol)) = 0 Then
                    buysellinfo(r, iOrderAmountCol) = CLng(Math.Truncate(CType(buysellinfo(r, iOrderMoneyCol), Decimal) / CType(stockinfo(r, isWriteCol), Decimal)))
                End If
            End If


            ''주문 넣기
            Account_number = CInt(objv.Range("Selected_Account").Value2)

            Call oShin_buysell.SetQueryName("SABA101U1")
            Call oShin_buysell.SetSingleData(0, CStr(objv.Range("Account_number_" & Account_number).Value2))
            Call oShin_buysell.SetSingleData(1, "redacted")
            Call oShin_buysell.SetSingleData(2, "redacted")
            Call oShin_buysell.SetSingleData(3, vbNullString)
            Call oShin_buysell.SetSingleData(4, vbNullString)
            Call oShin_buysell.SetSingleData(5, "0")
            Call oShin_buysell.SetSingleData(6, "00")
            Call oShin_buysell.SetSingleData(7, CStr(buyorsell))
            Call oShin_buysell.SetSingleData(8, "A" & CStr(stockinfo(r, isReadCol)))
            Call oShin_buysell.SetSingleData(9, CStr(buysellinfo(r, iOrderAmountCol)))
            If CLng(price) = PRICE_MARKET_PRICE Then Call oShin_buysell.SetSingleData(10, "0") Else Call oShin_buysell.SetSingleData(10, price)
            Call oShin_buysell.SetSingleData(11, "1")
            If CLng(price) = PRICE_MARKET_PRICE Then Call oShin_buysell.SetSingleData(12, "1") Else Call oShin_buysell.SetSingleData(12, "2")
            Call oShin_buysell.SetSingleData(13, "0")
            Call oShin_buysell.SetSingleData(14, "0")
            Call oShin_buysell.SetSingleData(15, vbNullString)
            Call oShin_buysell.SetSingleData(16, vbNullString)
            Call oShin_buysell.SetSingleData(21, "Y")

            ''매수매도
            Dim shit As Short = oShin_buysell.RequestData
            If shit <= 0 Then
                Handle_Error("[ERROR] 매수/매도 송신오류 : " & oShin_buysell.GetErrorState() & ", " & CStr(oShin_buysell.GetErrorMessage()) & ", " _
                & "buysell = " & CStr(buyorsell) & "종목코드 = " & CStr(stockinfo(r, isReadCol)) & ", 가격 = " & price & ", 주 = " & CStr(buysellinfo(r, iOrderAmountCol)) & ", r = " & r)
                'MsgBox("구매/판매 송신오류 : " & oShin_buysell.GetErrorState() & Chr(10) & CStr(oShin_buysell.GetErrorMessage()) & Chr(10) _
                '& "buysell = " & CStr(buyorsell) & "종목코드 = " & CStr(stockinfo(r, isReadCol)) & ", 가격 = " & price & ", 주 = " & CStr(buysellinfo(r, iOrderAmountCol)) & ", r = " & r, CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
                Continue For
            End If
            '주문을 다 넣었으면 내용 집어넣기
            buysellinfo(r, iOrderResultCol) = shit
            buysellinfo(r, iOrderBoxCol) = False
            req_remaining(LOAD_ARRAY_BUYSELL) += 1
        Next r

        '다했으면 엑셀에 반영
        If req_remaining(LOAD_ARRAY_BUYSELL) > 0 Then
            Timer_Buysell.Interval = 3000
            If Timer_Buysell.Enabled = False Then Timer_Buysell.Start()
            Writearray(LOAD_ARRAY_BUYSELL)
        End If

    End Sub
    Private Sub OShin_buysell_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles oShin_buysell.ReceiveData

        'error handling
        req_remaining(LOAD_ARRAY_BUYSELL) -= 1
        Dim nErr As Long = oShin_buysell.GetErrorState()
        If nErr > 0 Then
            Handle_Error("[ERROR] 매수/매도 수신오류 : " & oShin_buysell.GetErrorState() & ", " & CStr(oShin_buysell.GetErrorMessage()))
            Exit Sub
        End If

        Dim r As Integer
        For r = 1 To buysellinfo.GetUpperBound(0) + 1
            If e.rqid = CShort(buysellinfo(r, iOrderResultCol)) Then Exit For
        Next r
        If r = buysellinfo.GetUpperBound(0) + 1 Then
            Handle_Error("[ERROR] 매수/매도 거래내역을 찾을 수 없음 : " & e.rqid)
            Exit Sub
        End If

        nErr = CLng(oShin_buysell.GetSingleData(0))
        If nErr = 0 Then
            Handle_Error("[CRITICAL] 매수/매도 계좌주문 오류 : " & r & "번, 부족금액/수량 : " & CStr(oShin_buysell.GetSingleData(3)) & ", 가능수량 : " _
                & CStr(oShin_buysell.GetSingleData(4)) & ", 가능금액 : " & CStr(oShin_buysell.GetSingleData(5)))
            buysellinfo(r, iOrderResultCol) = "오류"
            Exit Sub
        End If
        '주문번호를 입력
        buysellinfo(r, iOrderResultCol) = CStr(oShin_buysell.GetSingleData(0))
        'Writearray(LOAD_ARRAY_BUYSELL)

    End Sub


    Public Sub Realtime_Stop_Handler(params As Object)

        Dim i = CInt(params)
        Select Case i
            Case START_OR_STOP_ALL
                Dim all As CRT_And_Autodeal
                For Each all In REALTIME_STORAGE
                    If all.Enabled = True Then all.RT_Stop()
                Next all
                Invoke(New WriteGrid(AddressOf UpdateGrid), New Object() {-1})
                Timer_Realtime.Stop()
            Case Else
                If REALTIME_STORAGE(i).Enabled = True Then
                    REALTIME_STORAGE(i).RT_Stop()
                    Invoke(New WriteGrid(AddressOf UpdateGrid), New Object() {i + EXCEL_REALTIME_CONVERSION_INDEX})
                End If
        End Select

    End Sub

    Public Sub Realtime_Start_Handler(params As Object)

        'Loadarray(LOAD_ARRAY_STOCKINFO) 'no needed, skip
        Loadarray(LOAD_ARRAY_AUTODEAL)
        Loadarray(LOAD_ARRAY_REALTIME)

        Dim i = CInt(params)
        Select Case i
            Case START_OR_STOP_ALL
                For iter As Integer = 0 To REALTIME_STORAGE.GetUpperBound(0) - 1
                    Realtime_start(iter)
                Next iter
            Case Else : Realtime_start(i)
        End Select

    End Sub

    Public Sub Realtime_start(idx As Integer)

        If CBool(realtime_checkbox(idx + 1, 1)) = False Then Exit Sub

        ' check if all the paramaters are valid
        Dim adr As CRT_And_Autodeal.Posinfo, sCode As String
        adr.rows = idx + (RealtimeStart - StockReadStart + 1)
        sCode = CStr(stockinfo(adr.rows, isReadCol))

        If sCode = vbNullString Or Len(sCode) <> 6 Then Exit Sub

        'if valid then assign, stop and restart
        adr.col_amount = isAmountCol
        adr.col_name = isNameCol
        adr.col_price = isWriteCol
        If REALTIME_STORAGE(idx).Enabled = True Then
            REALTIME_STORAGE(idx).RT_Stop()
            'Thread.Sleep(50)
        End If
        REALTIME_STORAGE(idx).RT_Start(sCode, adr)
        Invoke(New WriteGrid(AddressOf UpdateGrid), New Object() {adr.rows})
        If Timer_Realtime.Enabled = False Then
            Timer_Realtime.Interval = 400
            Timer_Realtime.Start()
        End If

    End Sub

    Public Sub GetBalanceInfo()

        Dim ret As Integer
        Account_number = CInt(objv.Range("Selected_Account").Value2)

        Call oShin_balance.SetQueryName("SABA609Q1")
        Call oShin_balance.SetSingleData(0, CStr(objv.Range("Account_number_" & Account_number).Value2))
        Call oShin_balance.SetSingleData(1, "redacted")
        Call oShin_balance.SetSingleData(2, "redacted")
        Call oShin_balance.SetSingleData(3, "1")
        Call oShin_balance.SetSingleData(4, "2")
        Call oShin_balance.SetSingleData(5, "0")
        Call oShin_balance.SetSingleData(6, "0")
        ret = oShin_balance.RequestData

        If ret <= 0 Then
            Handle_Error("[ERROR] 잔고조회 송신오류 : " & oShin_balance.GetErrorState() & ", " & CStr(oShin_balance.GetErrorMessage()))
            'MsgBox("잔고조회 송신오류 : " & oShin_balance.GetErrorState() & Chr(10) & CStr(oShin_balance.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

    End Sub

    Private Sub OShin_balance_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles oShin_balance.ReceiveData

        'error handling
        Dim nErr As Integer = oShin_balance.GetErrorState()
        If nErr > 0 Then
            Handle_Error("[ERROR] 잔고조회 수신오류 : " & oShin_balance.GetErrorState() & ", " & CStr(oShin_balance.GetErrorMessage()))
            'MsgBox("[balance] 수신오류 : " & oShin_balance.GetErrorState() & Chr(10) & CStr(oShin_balance.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        Loadarray(LOAD_ARRAY_BALANCE)

        For rows As Integer = 1 To favoriteinfo.GetUpperBound(0)
            favoriteinfo(rows, iFGoodPriceCol) = vbNullString
        Next rows

        For rows As Integer = 1 To autodealinfo.GetUpperBound(0)
            autodealinfo(rows, iBalGoodPriceCol) = vbNullString
            autodealinfo(rows, iBalOwnAmountCol) = vbNullString
        Next rows

        '잔고입력
        Account_Balance = CLng(oShin_balance.GetSingleData(12))

        Dim cnt As Short, compareCode As String, goodprice As Long
        Dim r As Integer, col As Integer, stockcode As String

        For cnt = 0 To oShin_balance.GetMultiRowCount - CShort(1)

            goodprice = CLng(oShin_balance.GetMultiData(cnt, 12)) '12 : 제비용단가, 28 : 단축코드, 10 : 주문가능수량
            compareCode = CStr(oShin_balance.GetMultiData(cnt, 28)) '12 : 제비용단가, 28 : 단축코드, 10 : 주문가능수량

            'Load stock first
            col = isReadCol
            For r = 1 To stockinfo.GetUpperBound(0)
                stockcode = CStr(stockinfo(r, col))
                If Len(stockcode) = 6 And StrComp(stockcode, compareCode, vbBinaryCompare) = 0 Then 'r, column
                    autodealinfo(r, iBalGoodPriceCol) = goodprice
                    autodealinfo(r, iBalOwnAmountCol) = CLng(oShin_balance.GetMultiData(cnt, 10)) '주문가능수량
                End If
            Next r

            'Load Favorite
            col = iFReadCol
            For r = 1 To favoriteinfo.GetUpperBound(0)
                stockcode = CStr(favoriteinfo(r, col))
                If Len(stockcode) = 6 And StrComp(stockcode, compareCode, vbBinaryCompare) = 0 Then   'r, column
                    favoriteinfo(r, iFGoodPriceCol) = goodprice
                End If
            Next r
        Next

        'write shit
        Dim thread As New Thread(AddressOf Thread_Writeshit)
        thread.Start()

    End Sub

    Public Sub Thread_Writeshit()
        objv.Range("Balance_" & Account_number).FormulaR1C1 = Account_Balance
        Writearray(LOAD_ARRAY_FAVORITE)
        Writearray(LOAD_ARRAY_BALANCE)
    End Sub
    Public Sub GetStockInfo()

        loadarray(LOAD_ARRAY_STOCKINFO)
        loadarray(LOAD_ARRAY_FAVORITE)

        Dim sCode As String
        req_remaining(LOAD_ARRAY_STOCKINFO) = 0
        req_remaining(LOAD_ARRAY_FAVORITE) = 0

        timer_tick_count(LOAD_ARRAY_STOCKINFO) = 0
        Timer_StockInfo.Interval = 100
        Timer_StockInfo.Enabled = True
        Timer_StockInfo.Start()

        For rs As Integer = 1 To stockinfo.GetUpperBound(0)
            stockinfo(rs, isNameCol) = vbNullString
            stockinfo(rs, isWriteCol) = vbNullString
            stockinfo(rs, isAmountCol) = vbNullString
            stockinfo(rs, iFNameCol) = vbNullString
            stockinfo(rs, iFWriteCol) = vbNullString
        Next rs

        For rs As Integer = 1 To stockinfo.GetUpperBound(0)
            'Parallel.For(1, stockinfo.GetUpperBound(0),
            'Sub(rs)

            sCode = CStr(stockinfo(rs, isReadCol))
            If sCode <> vbNullString Or Strings.Len(sCode) = 6 Then
                'Get stock price and amount
                oShin_PriceAmount.SetQueryName("SC")
                oShin_PriceAmount.SetSingleData(0, CStr(sCode))
                stockinfo(rs, isWriteCol) = oShin_PriceAmount.RequestData()
                If CShort(stockinfo(rs, isWriteCol)) <= 0 Then
                    Handle_Error("[ERROR] 현재가 송신오류(가격/거래량) : " & oShin_PriceAmount.GetErrorState() & ", " & CStr(oShin_PriceAmount.GetErrorMessage()))
                    'MsgBox("현재가 송신오류 : " & oShin_PriceAmount.GetErrorState() & Chr(10) & CStr(oShin_PriceAmount.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "suckman")
                    Exit Sub
                End If
                req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) + 1

                'and stock name
                oShin_name.SetQueryName("SB")
                oShin_name.SetSingleData(0, CStr(sCode))
                stockinfo(rs, isNameCol) = oShin_name.RequestData()
                If CShort(stockinfo(rs, isNameCol)) <= 0 Then
                    Handle_Error("[ERROR] 현재가 송신오류(이름) : " & oShin_name.GetErrorState() & ", " & CStr(oShin_name.GetErrorMessage()))
                    'MsgBox("현재가 송신오류 : " & oShin_name.GetErrorState() & Chr(10) & CStr(oShin_name.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
                    Exit Sub
                End If
                req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) + 1

            End If

            sCode = CStr(favoriteinfo(rs, iFReadCol))
            If sCode <> vbNullString Or Strings.Len(sCode) = 6 Then
                'Get stock price (only)
                oShin_Priceonly.SetQueryName("SC")
                oShin_Priceonly.SetSingleData(0, CStr(sCode))
                favoriteinfo(rs, iFWriteCol) = oShin_Priceonly.RequestData()
                If CShort(favoriteinfo(rs, iFWriteCol)) <= 0 Then
                    Handle_Error("[ERROR] 관심종목 송신오류(가격) : " & oShin_Priceonly.GetErrorState() & ", " & CStr(oShin_Priceonly.GetErrorMessage()))
                    'MsgBox("관심종목 송신오류 : " & oShin_Priceonly.GetErrorState() & Chr(10) & CStr(oShin_Priceonly.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "suckman")
                    Exit Sub
                End If
                req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) + 1

                'and stock name
                oShin_favname.SetQueryName("SB")
                oShin_favname.SetSingleData(0, CStr(sCode))
                favoriteinfo(rs, iFNameCol) = oShin_favname.RequestData()
                If CShort(favoriteinfo(rs, iFNameCol)) <= 0 Then
                    Handle_Error("[ERROR] 관심종목 송신오류(이름) : " & oShin_favname.GetErrorState() & ", " & CStr(oShin_favname.GetErrorMessage()))
                    'MsgBox("관심종목 송신오류 : " & oShin_favname.GetErrorState() & Chr(10) & CStr(oShin_favname.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
                    Exit Sub
                End If
                req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) + 1

            End If
            'End Sub)
        Next rs

    End Sub
    Private Sub OShin_favname_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles oShin_favname.ReceiveData

        req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) - 1
        Dim nErr As Integer = oShin_favname.GetErrorState()
        If nErr > 0 Then
            Handle_Error("[ERROR] 관심종목 수신오류(이름) : " & oShin_favname.GetErrorState() & ", " & CStr(oShin_favname.GetErrorMessage()))
            MsgBox("[Favname] 수신오류 : " & oShin_favname.GetErrorState() & Chr(10) & CStr(oShin_favname.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        Parallel.For(1, favoriteinfo.GetUpperBound(0),
        Sub(rs, state)
            If StrComp(CStr(favoriteinfo(rs, isNameCol)), CStr(e.rqid)) = 0 Then
                favoriteinfo(rs, isNameCol) = oShin_favname.GetSingleData(5)
                state.Break()
            End If
        End Sub)

    End Sub
    Private Sub OShin_Priceonly_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles oShin_Priceonly.ReceiveData

        req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) - 1
        Dim nErr As Integer = oShin_Priceonly.GetErrorState()
        If nErr > 0 Then
            Handle_Error("[ERROR] 관심종목 수신오류(가격) : " & oShin_Priceonly.GetErrorState() & ", " & CStr(oShin_Priceonly.GetErrorMessage()))
            'MsgBox("[Priceonly] 수신오류 : " & oShin_Priceonly.GetErrorState() & Chr(10) & CStr(oShin_Priceonly.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        Parallel.For(1, stockinfo.GetUpperBound(0),
        Sub(rs, state)
            If StrComp(CStr(favoriteinfo(rs, iFWriteCol)), CStr(e.rqid)) = 0 Then
                favoriteinfo(rs, iFWriteCol) = oShin_Priceonly.GetSingleData(3)
                state.Break()
            End If
        End Sub)

    End Sub

    Private Sub OShin_name_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles oShin_name.ReceiveData

        req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) - 1
        Dim nErr As Integer = oShin_name.GetErrorState()
        If nErr > 0 Then
            Handle_Error("[ERROR] 현재가 수신오류(이름) : " & oShin_name.GetErrorState() & ", " & CStr(oShin_name.GetErrorMessage()))
            'MsgBox("[stockname] 수신오류 : " & oShin_name.GetErrorState() & Chr(10) & CStr(oShin_name.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        Parallel.For(1, stockinfo.GetUpperBound(0),
        Sub(rs, state)
            If StrComp(CStr(stockinfo(rs, isNameCol)), CStr(e.rqid)) = 0 Then
                stockinfo(rs, isNameCol) = oShin_name.GetSingleData(5)
                state.Break()
            End If
        End Sub)

    End Sub

    Private Sub OShin_PriceAmount_ReceiveData(sender As Object, e As AxGIEXPERTCONTROLLib._DGiExpertControlEvents_ReceiveDataEvent) Handles oShin_PriceAmount.ReceiveData

        req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) - 1
        Dim nErr As Integer = oShin_PriceAmount.GetErrorState()
        If nErr > 0 Then
            Handle_Error("[ERROR] 현재가 수신오류(가격,거래량) : " & oShin_PriceAmount.GetErrorState() & ", " & CStr(oShin_PriceAmount.GetErrorMessage()))
            'MsgBox("[PriceAmount] 수신오류 : " & oShin_PriceAmount.GetErrorState() & Chr(10) & CStr(oShin_PriceAmount.GetErrorMessage()), CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        Parallel.For(1, stockinfo.GetUpperBound(0),
        Sub(rs, state)
            If StrComp(CStr(stockinfo(rs, isWriteCol)), CStr(e.rqid)) = 0 Then
                stockinfo(rs, isWriteCol) = oShin_PriceAmount.GetSingleData(3)
                stockinfo(rs, isAmountCol) = oShin_PriceAmount.GetSingleData(7).ToString
                state.Break()
            End If
        End Sub)

    End Sub

    'Delegate Sub WriteStockinfo(ByRef rs As Integer, ByRef cols As Integer, ByRef value As Object)
    Public Sub Assign_stockinfo(ByRef rs As Integer, ByRef cols As Integer, ByRef value As Object)
        stockinfo(rs, cols) = value
    End Sub

    Delegate Sub Write_Onbehalf(ByRef val As Integer)
    Public Sub Writearray(ByRef val As Integer)

        Dim tmp(,) As Object
        Try
            Select Case val
                Case LOAD_ARRAY_STOCKINFO
                    tmp = CType(objMain.Range(objMain.Cells(StockReadStart, StockReadColumn + 1), objMain.Cells(StockReadFinish, TradeAmountColumn)).Value2, Object(,))
                    For c As Integer = 1 To tmp.GetUpperBound(1)
                        For r As Integer = 1 To tmp.GetUpperBound(0)
                            tmp(r, c) = stockinfo(r, c + 1)
                        Next
                    Next
                    objMain.Range(objMain.Cells(StockReadStart, StockReadColumn + 1), objMain.Cells(StockReadFinish, TradeAmountColumn)).FormulaR1C1 = tmp
                Case LOAD_ARRAY_FAVORITE
                    objMain.Range(objMain.Cells(FavoriteStart, FavoriteReadColumn), objMain.Cells(FavoriteFinish, Goodprice_deal_Favorite_Column)).FormulaR1C1 = favoriteinfo
                Case LOAD_ARRAY_BALANCE
                    Writearray(LOAD_ARRAY_FAVORITE)
                    tmp = CType(objMain.Range(objMain.Cells(StockReadStart, Autodeal_ownstock_col), objMain.Cells(StockReadFinish + 1, Goodprice_deal_AUTO_Column)).Value2, Object(,))
                    For c As Integer = 1 To tmp.GetUpperBound(1)
                        For r As Integer = 1 To tmp.GetUpperBound(0)
                            tmp(r, c) = autodealinfo(r, c + 1)
                        Next r
                    Next c
                    objMain.Range(objMain.Cells(StockReadStart, Autodeal_ownstock_col), objMain.Cells(StockReadFinish, Goodprice_deal_AUTO_Column)).FormulaR1C1 = tmp
                Case LOAD_ARRAY_AUTODEAL
                    tmp = CType(objMain.Range(objMain.Cells(StockReadStart, Autodeal_status_col), objMain.Cells(StockReadFinish, Autodeal_status_col)).Value2, Object(,))
                    For c As Integer = 1 To tmp.GetUpperBound(1)
                        For r As Integer = 1 To tmp.GetUpperBound(0)
                            tmp(r, c) = autodealinfo(r, iAutoStatusCol)
                        Next
                    Next
                    objMain.Range(objMain.Cells(StockReadStart, Autodeal_status_col), objMain.Cells(StockReadFinish, Autodeal_status_col)).FormulaR1C1 = tmp
                Case LOAD_ARRAY_BUYSELL
                    objMain.Range(objMain.Cells(StockReadStart, Order_Ckbox_Column), objMain.Cells(StockReadFinish, Order_result_Column)).FormulaR1C1 = buysellinfo
            End Select
        Catch ex As Exception
            Handle_Error("[WARNING] WriteArray 도중 오류가 발생헀습니다. 메시지 : " & ex.Message)
        End Try
    End Sub

    Private Sub Loadarray(ByRef val As Integer)

        'Dim usedRange As Range
        'usedRange = objMain.Range().Address
        'stockinfo = CType(usedRange.Value2, Object(,)) '
        Try
            Select Case val
                Case LOAD_ARRAY_ALL
                    stockinfo = CType(objMain.Range(objMain.Cells(StockReadStart, StockReadColumn), objMain.Cells(StockReadFinish + 1, TradeAmountColumn)).Value2, Object(,))
                    favoriteinfo = CType(objMain.Range(objMain.Cells(FavoriteStart, FavoriteReadColumn), objMain.Cells(FavoriteFinish + 1, Goodprice_deal_Favorite_Column)).Value2, Object(,))
                    'balanceinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Autodeal_ownstock_col), objMain.Cells(StockReadFinish + 1, Goodprice_deal_AUTO_Column)).Value2, Object(,))
                    autodealinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Autodeal_CheckBox_Column), objMain.Cells(StockReadFinish + 1, Autodeal_status_col)).Value2, Object(,))
                    buysellinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Order_Ckbox_Column), objMain.Cells(StockReadFinish + 1, Order_result_Column)).Value2, Object(,))
                    realtime_checkbox = CType(objMain.Range(objMain.Cells(RealtimeStart, realtime_checkbox_column), objMain.Cells(RealtimeFinish + 1, realtime_checkbox_column)).Value2, Object(,))
                    doctrineinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Doctrine_Number_Column), objMain.Cells(StockReadFinish + 1, Doctrine_Arg2_Column)).Value2, Object(,))
                Case LOAD_ARRAY_STOCKINFO
                    stockinfo = CType(objMain.Range(objMain.Cells(StockReadStart, StockReadColumn), objMain.Cells(StockReadFinish + 1, TradeAmountColumn)).Value2, Object(,))
                Case LOAD_ARRAY_FAVORITE
                    favoriteinfo = CType(objMain.Range(objMain.Cells(FavoriteStart, FavoriteReadColumn), objMain.Cells(FavoriteFinish + 1, Goodprice_deal_Favorite_Column)).Value2, Object(,))
                Case LOAD_ARRAY_BALANCE
                    Loadarray(LOAD_ARRAY_FAVORITE)
                    Loadarray(LOAD_ARRAY_AUTODEAL)
                Case LOAD_ARRAY_AUTODEAL
                    Loadarray(LOAD_ARRAY_STOCKINFO)
                    Loadarray(LOAD_ARRAY_DOCTRINES)
                    autodealinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Autodeal_CheckBox_Column), objMain.Cells(StockReadFinish + 1, Autodeal_status_col)).Value2, Object(,))
                Case LOAD_ARRAY_BUYSELL
                    Loadarray(LOAD_ARRAY_STOCKINFO)
                    buysellinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Order_Ckbox_Column), objMain.Cells(StockReadFinish + 1, Order_result_Column)).Value2, Object(,))
                Case LOAD_ARRAY_REALTIME
                    realtime_checkbox = CType(objMain.Range(objMain.Cells(RealtimeStart, realtime_checkbox_column), objMain.Cells(RealtimeFinish + 1, realtime_checkbox_column)).Value2, Object(,))
                Case LOAD_ARRAY_DOCTRINES
                    doctrineinfo = CType(objMain.Range(objMain.Cells(StockReadStart, Doctrine_Number_Column), objMain.Cells(StockReadFinish + 1, Doctrine_Arg2_Column)).Value2, Object(,))
            End Select
        Catch ex As Exception
            Handle_Error("[WARNING] ReadArray 도중 오류가 발생헀습니다. 메시지 : " & ex.Message)
        End Try

    End Sub

    Private Sub Load_settings()

        AUTODEAL_ENABLED_MASTERKEY = CBool(objv.Range("Autodeal_Enabled").Value2)
        AccountNumString = CStr(objv.Range("Account_number_" & CStr(objv.Range("Selected_Account").Value2)).Value2)

        'favorite read
        FavoriteReadColumn = GetColumn("FavoriteStart")
        FavoriteStart = Getrow("FavoriteStart")
        FavoriteFinish = Getrow("FavoriteFinish")
        FavoriteWriteColumn = GetColumn("Favorite_target_column")
        FavoriteNameColumn = GetColumn("FavoriteNameColumn")

        'stocklist read
        StockReadColumn = GetColumn("StockreadStart")
        StockReadStart = Getrow("StockreadStart")
        StockReadFinish = Getrow("StockreadFinish")
        StockWriteColumn = GetColumn("Stockread_target_column")
        StockNameColumn = GetColumn("StockNameColumn")

        'real code read
        RealtimeReadColumn = GetColumn("RealtimeStart")
        RealtimeStart = Getrow("RealtimeStart")
        RealtimeFinish = Getrow("RealtimeFinish")
        RealtimeWritecolumn = GetColumn("Realtime_target_column")
        realtime_checkbox_column = GetColumnInMultiRange("realtime_checkbox_range")

        TradeAmountColumn = GetColumn("TradeAmountColumn")

        '제비용단가 위치 구하기
        Goodprice_deal_AUTO_Column = GetColumn("goodprice_deal_AUTO")
        Goodprice_deal_Favorite_Column = GetColumn("goodprice_deal_Favorite")

        '자동거래용 칸구하기
        Autodeal_CheckBox_Column = GetColumnInMultiRange("SelectALL_AUTOdeal")
        Autodeal_badsell_Price_Col = GetColumn("Autodeal_badsell_Price_Col")
        Autodeal_goodsell_Price_Col = GetColumn("Autodeal_goodsell_Price_Col")
        Autodeal_goodpercent_Col = GetColumn("Autodeal_goodpercent_Col")
        Autodeal_badpercent_Col = GetColumn("Autodeal_badpercent_Col")
        Autodeal_ownstock_col = GetColumn("Autodeal_ownstock_col")
        Autodeal_status_col = GetColumn("Autodeal_status_col")

        '매수/매도/관심종목란
        Order_Ckbox_Column = GetColumnInMultiRange("SelectALL_deal")
        Order_price_Column = GetColumn("Order_price_Column") '가격column
        Order_amount_Column = GetColumn("Order_amount_Column") '주column
        Order_money_Column = GetColumn("Order_money_Column") '금액column
        Order_result_Column = GetColumn("Order_result_Column") '결과column

        Doctrine_Number_Column = GetColumn("Doctrine_Number_Column")
        Doctrine_Arg1_Column = GetColumn("Doctrine_Arg1_Column")
        Doctrine_Arg2_Column = GetColumn("Doctrine_Arg2_Column")

        loadarray(LOAD_ARRAY_ALL)

        ReDim REALTIME_STORAGE(RealtimeFinish - RealtimeStart + 1)
        For i As Integer = 0 To REALTIME_STORAGE.GetUpperBound(0) ' - 1
            REALTIME_STORAGE(i) = New CRT_And_Autodeal()
        Next

        '일단은 하드코딩으로 제작
        For i As Integer = 0 To REALTIME_STORAGE.GetUpperBound(0) - 1
            DataGridView1.Rows.Add({"", "", False, "", "", "", "", "", "", "", "", "", ""})
        Next

        Dim col_offset As Integer
        col_offset = TradeAmountColumn - stockinfo.GetUpperBound(1)
        isReadCol = StockReadColumn - col_offset
        isNameCol = StockNameColumn - col_offset
        isWriteCol = StockWriteColumn - col_offset
        isAmountCol = TradeAmountColumn - col_offset

        col_offset = Goodprice_deal_Favorite_Column - favoriteinfo.GetUpperBound(1)
        iFReadCol = FavoriteReadColumn - col_offset
        iFWriteCol = FavoriteWriteColumn - col_offset
        iFNameCol = FavoriteNameColumn - col_offset
        iFGoodPriceCol = Goodprice_deal_Favorite_Column - col_offset

        col_offset = Autodeal_status_col - autodealinfo.GetUpperBound(1)
        iAutoBoxCol = Autodeal_CheckBox_Column - col_offset
        iAutoBadsellCol = Autodeal_badsell_Price_Col - col_offset
        iAutoGoodsellCol = Autodeal_goodsell_Price_Col - col_offset
        iAutoGoodPerCol = Autodeal_goodpercent_Col - col_offset
        iAutoBadPerCol = Autodeal_badpercent_Col - col_offset
        iAutoOwnCol = Autodeal_ownstock_col - col_offset
        iAutoStatusCol = Autodeal_status_col - col_offset
        iBalGoodPriceCol = Goodprice_deal_AUTO_Column - col_offset
        iBalOwnAmountCol = Autodeal_ownstock_col - col_offset

        col_offset = Order_result_Column - buysellinfo.GetUpperBound(1)
        iOrderBoxCol = Order_Ckbox_Column - col_offset
        iOrderPriceCol = Order_price_Column - col_offset
        iOrderAmountCol = Order_amount_Column - col_offset
        iOrderMoneyCol = Order_money_Column - col_offset
        iOrderResultCol = Order_result_Column - col_offset

        col_offset = Doctrine_Arg2_Column - doctrineinfo.GetUpperBound(1)
        iDocNumCol = Doctrine_Number_Column - col_offset
        iDocArg1Col = Doctrine_Arg1_Column - col_offset
        iDocArg2Col = Doctrine_Arg2_Column - col_offset

        EXCEL_REALTIME_CONVERSION_INDEX = (RealtimeStart - StockReadStart) + 1

    End Sub

    Public Function IsAutoDealEnabled() As Boolean
        IsAutoDealEnabled = AUTODEAL_ENABLED_MASTERKEY
    End Function

    Private Function Getrow(ByRef sKey As String) As Integer

        Dim tmp As String
        tmp = Trim(CStr(objv.Range(sKey).Formula))
        tmp = objMain.Range(Strings.Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare))).Address(ReferenceStyle:=XlReferenceStyle.xlR1C1)
        Getrow = CInt(Mid(tmp, 2, InStr(1, tmp, "C", vbTextCompare) - 2))  'row 구하기

    End Function

    Private Function GetColumn(ByRef sKey As String) As Integer

        Dim tmp As String
        tmp = Trim(CStr(objv.Range(sKey).Formula))
        tmp = objMain.Range(Strings.Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare))).Address(ReferenceStyle:=XlReferenceStyle.xlR1C1)
        GetColumn = CInt(Strings.Right(tmp, Len(tmp) - InStr(1, tmp, "C", vbTextCompare))) 'column 구하기

    End Function

    Private Function GetColumnInMultiRange(ByRef sKey As String) As Integer

        Dim tmp As String
        tmp = Trim(CStr(objv.Range(sKey).Formula))
        tmp = objMain.Range(Strings.Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare))).Address(ReferenceStyle:=XlReferenceStyle.xlR1C1)
        GetColumnInMultiRange = CInt(Mid(tmp, InStr(1, tmp, "C", vbTextCompare) + 1, InStr(1, tmp, ":", vbTextCompare) - InStr(1, tmp, "C", vbTextCompare) - 1)) '범위에서 column 구하기

    End Function


    '========================================================
    '   디버그용 함수들 (Write_finished, Errormessage) for debug
    '      * 다른 쓰레드에서 접근하기 위한 함수
    '      * Label1.text에 값을 기록하여 디버그에 활용
    '========================================================
    Private Sub Label1_TextChanged(sender As Object, e As EventArgs) Handles Label1.TextChanged
        Label1.SelectionStart = Label1.Text.Length
    End Sub

    Public Sub Handle_Error(sss As String)
        Try
            Me.Invoke(New WriteErrMsg(AddressOf Errormessage), New Object() {sss})
        Catch exe As Exception
            If BYEBYE = False Then MsgBox("byebye error : " & exe.Message & Chr(10) & "Message: " & sss)
        End Try
    End Sub

    Delegate Sub WriteErrMsg(ByVal str As String)
    Private Sub Errormessage(ByVal str As String)
        Label1.Text = str & vbCrLf & Label1.Text
    End Sub

    Public Delegate Sub WriteGrid(ByVal rows As Integer)
    Public Sub UpdateGrid(ByVal rows As Integer)

        If AUTODEAL_ENABLED_MASTERKEY Then AutoStatus.Text = "Enabled" Else AutoStatus.Text = "Disabled"
        Select Case AUTODEAL_ENABLED_MASTERKEY
            Case True
                AutoStatus.ForeColor = Color.Blue
            Case False
                AutoStatus.ForeColor = Color.Red
        End Select

        Dim d_row As Integer, str As String
        If rows = -1 Then
            For d_row = RealtimeStart To RealtimeFinish
                rows = d_row - StockReadStart + 1
                If REALTIME_STORAGE(d_row - RealtimeStart).Enabled = True Then str = "●" Else str = ""
                DataGridView1.Rows(d_row - RealtimeStart).SetValues(New Object() {str, autodealinfo(rows, iAutoStatusCol), autodealinfo(rows, iAutoBoxCol), stockinfo(rows, isReadCol), stockinfo(rows, isNameCol), stockinfo(rows, isWriteCol), stockinfo(rows, isAmountCol),
                                            autodealinfo(rows, iBalOwnAmountCol), autodealinfo(rows, iBalGoodPriceCol), autodealinfo(rows, iAutoBadsellCol), autodealinfo(rows, iAutoBadPerCol), autodealinfo(rows, iAutoGoodsellCol), autodealinfo(rows, iAutoGoodPerCol)})
            Next d_row
            Exit Sub
        End If
        d_row = rows - EXCEL_REALTIME_CONVERSION_INDEX 'rows는 배열로 나오기 때문에 그만큼 밑으로 빼줘야 최신화가 가능
        If REALTIME_STORAGE(d_row).Enabled = True Then str = "●" Else str = ""
        DataGridView1.Rows(d_row).SetValues(New Object() {str, autodealinfo(rows, iAutoStatusCol), autodealinfo(rows, iAutoBoxCol), stockinfo(rows, isReadCol), stockinfo(rows, isNameCol), stockinfo(rows, isWriteCol), stockinfo(rows, isAmountCol),
                                            autodealinfo(rows, iBalOwnAmountCol), autodealinfo(rows, iBalGoodPriceCol), autodealinfo(rows, iAutoBadsellCol), autodealinfo(rows, iAutoBadPerCol), autodealinfo(rows, iAutoGoodsellCol), autodealinfo(rows, iAutoGoodPerCol)})
    End Sub

End Class