VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public WithEvents oShin_PriceAmount As GIExpertControl64
Attribute oShin_PriceAmount.VB_VarHelpID = -1
Public WithEvents oShin_PriceOnly As GIExpertControl64
Attribute oShin_PriceOnly.VB_VarHelpID = -1
Public WithEvents oShin_name As GIExpertControl64
Attribute oShin_name.VB_VarHelpID = -1
Public WithEvents oShin_favname As GIExpertControl64
Attribute oShin_favname.VB_VarHelpID = -1
'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ==========================================
' 제공되는 함수 (legacy)
' ==========================================
' GetCurrentPrice(종목코드, 셀주소(row), 셀주소(column))                                (종목코드 현재가 불러오기)
' GetRealtimePrice(종목코드, 셀주소(row), 셀주소(column), 실시간코드:0~4)               (종목코드 현재가 실시간으로 불러오기)
' GetStockName(종목코드, 셀주소(row), 셀주소(column))                                   (종목코드 이름 불러오기)
' GetTradeAmount(종목코드, 셀주소(row), 셀주소(column))                                 (종목코드 거래량 불러오기)
' StopRealtimePrice(실시간코드)                                                         (종목코드 현재가 실시간 불러오기 중단)
' StopALL()                                                                             (모든실시간 불러오기 중단)
' GetCurrentBalance()                                                                   (저장된 1번과 2번의 D+2 예수금을 업데이트함)
' #GetMyStock(종목코드, 셀주소)                                                         (나의 현재 보유주식 확인)
' SellStock(종목코드, 가격=-1:시장가, 주, 셀주소(row), 셀주소(column), 자동거래=0:일반주문, -1:손절매도, 1:익절매도)
' BuyStock(종목코드, 가격=-1:시장가, 주, 셀주소(row), 셀주소(column), 자동거래=0:일반주문, -1:손절매도, 1:익절매도)

' GetRow(ByRef sKey As String) As Integer
' GetColumn(ByRef sKey As String) As Integer
' GetColumnInMultiRange(ByRef sKey As String) As Integer
' GetRange(ByRef sKey As String) As String
' SetSettings(ByRef sKey As String, Value, indirect As Boolean)
' GetSettings(ByRef sKey As String) As Variant
' UsePlugin()

'=========================
' legacy implementation
'=========================
' 현재가,이름 : 1번(종목), 2번(0: ID, 1:갱신위치, 2: Type)
Const REQ_INVALID_ID = 0
Const REQ_ID = 0
Const REQ_WRITE_ROW = 1
Const REQ_WRITE_COL = 2
Const REQ_TYPE = 3 'Type 1 : 가격조회, 2 : 이름조회
    '요청 이름 (현재가 ,이름)
    Const REQ_TYPE_PRICE = 1
    Const REQ_TYPE_NAME = 2
    Const REQ_TYPE_TRADEAMOUNT = 3
    ' 실시간 : (가격, 거래량)
    Const REQ_TYPE_REALPRICE = 1
    Const REQ_TYPE_REALAMOUNT = 2
Const REQ_STOCKCODE = 4
Const REQ_PRICE = 5
Const REQ_AMOUNT = 6
Const REQ_AUTOSTOP = 7

Const RQSIZE = 256
Const RQRTSIZE = 20 + 1

Const CRAZYBUY_TIME_BUY = True
Const CRAZYBUY_TIME_SELL = False

'=========================
' legacy 엑셀 API 구현을 위한 저장소
'=========================
Dim RQ_ONCE_VARIABLE(RQSIZE, 3)
Dim RQ_CURPRICE_REALTIME(RQRTSIZE, 4)
Dim RQ_BUYSELL_VARIABLE(RQSIZE, 7)
'---------------------------------------------------------------------------------------------------------
'legacy monitoring events
Public WithEvents objShinhan As GIExpertControl64
Attribute objShinhan.VB_VarHelpID = -1
Public WithEvents objShinhan_balance As GIExpertControl64
Attribute objShinhan_balance.VB_VarHelpID = -1
Public WithEvents objShinhan_buysell As GIExpertControl64
Attribute objShinhan_buysell.VB_VarHelpID = -1
Public WithEvents objShinhan_RT As GIExpertControl64
Attribute objShinhan_RT.VB_VarHelpID = -1
Public WithEvents objShinhan_RT2 As GIExpertControl64
Attribute objShinhan_RT2.VB_VarHelpID = -1
Public WithEvents objShinhan_RT3 As GIExpertControl64
Attribute objShinhan_RT3.VB_VarHelpID = -1
Public WithEvents objShinhan_RT4 As GIExpertControl64
Attribute objShinhan_RT4.VB_VarHelpID = -1
Public WithEvents objShinhan_RT5 As GIExpertControl64
Attribute objShinhan_RT5.VB_VarHelpID = -1
Public WithEvents objShinhan_RT6 As GIExpertControl64
Attribute objShinhan_RT6.VB_VarHelpID = -1
Public WithEvents objShinhan_RT7 As GIExpertControl64
Attribute objShinhan_RT7.VB_VarHelpID = -1
Public WithEvents objShinhan_RT8 As GIExpertControl64
Attribute objShinhan_RT8.VB_VarHelpID = -1
Public WithEvents objShinhan_RT9 As GIExpertControl64
Attribute objShinhan_RT9.VB_VarHelpID = -1
Public WithEvents objShinhan_RT10 As GIExpertControl64
Attribute objShinhan_RT10.VB_VarHelpID = -1
Public WithEvents objShinhan_RT11 As GIExpertControl64
Attribute objShinhan_RT11.VB_VarHelpID = -1
Public WithEvents objShinhan_RT12 As GIExpertControl64
Attribute objShinhan_RT12.VB_VarHelpID = -1
Public WithEvents objShinhan_RT13 As GIExpertControl64
Attribute objShinhan_RT13.VB_VarHelpID = -1
Public WithEvents objShinhan_RT14 As GIExpertControl64
Attribute objShinhan_RT14.VB_VarHelpID = -1
Public WithEvents objShinhan_RT15 As GIExpertControl64
Attribute objShinhan_RT15.VB_VarHelpID = -1
Public WithEvents objShinhan_RT16 As GIExpertControl64
Attribute objShinhan_RT16.VB_VarHelpID = -1
Public WithEvents objShinhan_RT17 As GIExpertControl64
Attribute objShinhan_RT17.VB_VarHelpID = -1
Public WithEvents objShinhan_RT18 As GIExpertControl64
Attribute objShinhan_RT18.VB_VarHelpID = -1
Public WithEvents objShinhan_RT19 As GIExpertControl64
Attribute objShinhan_RT19.VB_VarHelpID = -1
Public WithEvents objShinhan_RT20 As GIExpertControl64
Attribute objShinhan_RT20.VB_VarHelpID = -1





'========================================================================================================
' 엑셀 / VB 공히 사용하는 환경 변수
'========================================================================================================
Const ORDER_BUY As Integer = 2              '구매/판매를 구분하기 위해 사용하는 변수. 실제 구매요청시에 동일한 값이 사용됨
Const ORDER_SELL As Integer = 1
Const PRICE_MARKET_PRICE = -1               '지정가와 시장가를 구분하기 위해 사용한 음의 변수
'---------------------------------------------------------------------------------------------------------
Const NOT_AUTOSTOP = 0                      '자동구매 여부를 판별하기 위한 환경변수, Buysell_Backend 함수에 활용됨
Const AUTOSTOP_GOOD = 1
Const AUTOSTOP_BAD = -1
'---------------------------------------------------------------------------------------------------------
Dim req_remaining(5) As Integer
Dim timer_tick_count(5) As Integer


'========================================================================================================
' 엑셀에서만 사용하는 환경 변수
'========================================================================================================
Const SET_DIRECT = False                    'SetSettings Function에서 변수를 찾아가서 수정할지 직접 수정할지를 나타내는 환경변수
Const SET_INDIRECT = True
'---------------------------------------------------------------------------------------------------------
Dim NOT_USE_PLUGIN As Boolean               '플러그인 사용여부를 저장해두는 변수
'---------------------------------------------------------------------------------------------------------



'========================================================================================================
' 엑셀에서 값을 불러와서 임시로 저장해두는 배열과 관련된 변수
'   (직접 Value2등으로 액세스 할 경우에 성능이 매우 떨어지기 때문에 별도의 저장소를 생성하였음)
'   (동일한 최적화를 VBA에 적용)
'   (LoadArray로 불러옴, WriteArray로 쓰기 / 기타 내용은 직접 엑세스)
'========================================================================================================
Const LOAD_ARRAY_ALL As Integer = 0         'LoadArray 함수에서 사용, 셀의 값을 모든 저장소에 로드
Const LOAD_ARRAY_STOCKINFO As Integer = 1   '단순조회 및 실시간조회에 사용되는 값만 로드 (저장소 : stockinfo)
Const LOAD_ARRAY_FAVORITE As Integer = 2    '관심종목과 관련된 셀들을 로드 (저장소 : favoriteinfo)
Const LOAD_ARRAY_BALANCE As Integer = 3     '현재 보유주식과 제비용단가 내용을 로드 (저장소 : autodealinfo)
'                                               - (본 함수 호출시 단순조회와 관심종목을 같이 불러옴)
Const LOAD_ARRAY_AUTODEAL As Integer = 4    '자동거래에 사용되는 셀들을 로드 (저장소 : autodealinfo)
'                                               - (본 함수 호출시 단순조회 셀도 같이 로드됨)
Const LOAD_ARRAY_BUYSELL As Integer = 5     '구매/판매에 사용되는 셀들을 로드 (저장소 : buysellinfo)
'                                               - (본 함수 호출시 단순조회 셀도 같이 로드됨)
Const LOAD_ARRAY_REALTIME As Integer = 6    '실시간 조회를 할지 말지 확인하는 값을 불러옴 (저장소 : stockinfo)
'                                               - (주의 : stockinfo와는 별도의 셀인 realtime_checkbox에 저장됨)
Const LOAD_ARRAY_DOCTRINES As Integer = 7   '독트린 관련 내용을 불러옴 (저장소 : doctrineinfo)
'---------------------------------------------------------------------------------------------------------
Dim stockinfo() As Variant                  '단순조회 및 실시간조회에 사용
Dim favoriteinfo() As Variant               '관심종목과 관련된 셀들을 로드
'Dim balanceinfo() As Variant                '현재 보유주식과 제비용단가 내용을 로드
Dim autodealinfo() As Variant               '자동거래에 사용되는 셀들을 로드
Dim buysellinfo() As Variant                '구매/판매에 사용되는 셀들을 로드
Dim realtime_checkbox() As Variant          '실시간 조회를 할지 말지 확인하는 값
'---------------------------------------------------------------------------------------------------------
Dim doctrineinfo() As Variant               '독트린 관련 내용 저장소

'========================================================================================================
' 임시로 저장해두는 배열과 실제 엑셀 변수 위치와의 괴리를 해결하기 위해 작성한 별도의 변수
'   (Load_settings에서 해당 값을 계산해줌, 실제 VB 계산시 해당 위치를 사용)
'   (행의 시작과 종료는 1 to 배열의 UpperBound를 사용하여 계산)
'========================================================================================================
Dim isReadCol As Integer                    '단순 1회성 조회용 종목번호 열 위치
Dim isWriteCol As Integer                   '단순조회 현재가 열 위치
Dim isNameCol As Integer                    '단순조회 종목이름 열 위치
Dim isAmountCol As Integer                  '실시간/단순조회 거래량 열 위치
'---------------------------------------------------------------------------------------------------------
Dim iFReadCol As Integer                    '관심종목 종목번호 열 위치
Dim iFNameCol As Integer                    '관심종목 종목이름 열 위치
Dim iFWriteCol As Integer                   '관심종목 현재가 열 위치
Dim iFGoodPriceCol As Integer               '잔고조회 / 제비용단가 : 관심종목란에 사용되는 열 위치
'---------------------------------------------------------------------------------------------------------
Dim iBalGoodPriceCol As Integer             '잔고조회 / 제비용단가 : 자동거래시에 사용되는 열 위치
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
Dim iAutoBoxCol As Integer                  '자동거래 사용여부 열 위치
Dim iAutoBadsellCol As Integer              '자동거래 손절가격 열 위치
Dim iAutoGoodsellCol As Integer             '자동거래 익절가격 열 위치
Dim iAutoGoodPerCol As Integer              '자동거래 익절율 계산용 열 위치 (제비용단가 기준)
Dim iAutoBadPerCol As Integer               '자동거래 손절율 계산용 열 위치 (제비용단가 기준)
Dim iAutoOwnCol As Integer                  '잔고조회 / 자동거래 보유주식 계산용 열 위치
Dim iAutoStatusCol As Integer               '자동거래 결과 기록용 열 위치


'========================================================================================================
'엑셀에서 시트를 변경하더라도 각 시트의 절대위치를 잡기 위해 선언한 변수
'   (VB.net 에서 직접 계산할 수는 없고, Load_settings 간 불러온 뒤 오프셋 확인용도로 사용)
'   (v 시트에 위치한 값을 불러오며, 행 / 열 추가시 변경해야할 수도 있음)
'========================================================================================================
Dim StockReadColumn As Integer      '단순 1회성 조회용 종목번호 열 위치
Dim StockWriteColumn As Integer     '단순조회 현재가 열 위치
Dim StockNameColumn As Integer      '단순조회 종목이름 열 위치
Dim StockReadStart As Integer       '단순조회 행 시작위치
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

Function GetRow(ByRef sKey As String) As Integer

    Dim tmp As String
    tmp = Trim(Sheets("v").Range(sKey).Formula)
    tmp = Sheets("Main").Range(Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare))).Address(ReferenceStyle:=xlR1C1)
    GetRow = CInt(Mid(tmp, 2, InStr(1, tmp, "C", vbTextCompare) - 2))  'row 구하기
    
End Function

Function GetColumn(ByRef sKey As String) As Integer

    Dim tmp As String
    tmp = Trim(Sheets("v").Range(sKey).Formula)
    tmp = Sheets("Main").Range(Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare))).Address(ReferenceStyle:=xlR1C1)
    GetColumn = CInt(Right(tmp, Len(tmp) - InStr(1, tmp, "C", vbTextCompare))) 'column 구하기
    
End Function

Function GetColumnInMultiRange(ByRef sKey As String) As Integer

    Dim tmp As String
    tmp = Trim(Sheets("v").Range(sKey).Formula)
    tmp = Sheets("Main").Range(Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare))).Address(ReferenceStyle:=xlR1C1)
    GetColumnInMultiRange = CInt(Mid(tmp, InStr(1, tmp, "C", vbTextCompare) + 1, InStr(1, tmp, ":", vbTextCompare) - InStr(1, tmp, "C", vbTextCompare) - 1)) '범위에서 column 구하기

End Function

Function GetRange(ByRef sKey As String) As String

    Dim tmp As String
    tmp = Trim(Sheets("v").Range(sKey).Formula)
    GetRange = Right(tmp, Len(tmp) - InStr(1, tmp, "!", vbTextCompare)) 'range 구하기
    
End Function

Sub SetSettings(ByRef sKey As String, Value, indirect As Boolean)
    
    Select Case indirect
        
        Case SET_INDIRECT: Sheets("Main").Range(GetRange(sKey)).FormulaR1C1 = Value
        Case SET_DIRECT: Sheets("v").Range(sKey).FormulaR1C1 = Value
        
    End Select
    
End Sub

Function GetSettings(ByRef sKey As String) As Variant
    
    GetSettings = Sheets("v").Range(sKey).Value2
    
End Function

Function UsePlugin() As Boolean
    UsePlugin = Not (NOT_USE_PLUGIN)
End Function

Function CheckRemaining(ByRef val As Integer) As Integer
    CheckRemaining = req_remaining(val)
End Function

Function CheckTimerCount(ByRef val As Integer) As Integer
    CheckTimerCount = timer_tick_count(val)
End Function

Function SetTimerCount(ByRef val As Integer, ByRef iSet As Integer)
    timer_tick_count(val) = iSet
End Function

Sub loadarray(ByRef val As Integer)

    Select Case val
        Case LOAD_ARRAY_ALL
            With Sheets("Main")
                stockinfo = .Range(.Cells(StockReadStart, StockReadColumn).Address, .Cells(StockReadFinish + 1, TradeAmountColumn).Address).Value2
                favoriteinfo = .Range(.Cells(FavoriteStart, FavoriteReadColumn).Address, .Cells(FavoriteFinish + 1, Goodprice_deal_Favorite_Column).Address).Value2
                'balanceinfo = .Range(.Cells(StockReadStart, Autodeal_ownstock_col).Address, .Cells(StockReadFinish + 1, Goodprice_deal_AUTO_Column).Address).Value2
                autodealinfo = .Range(.Cells(StockReadStart, Autodeal_CheckBox_Column).Address, .Cells(StockReadFinish + 1, Autodeal_status_col).Address).Value2
                buysellinfo = .Range(.Cells(StockReadStart, Order_Ckbox_Column).Address, .Cells(StockReadFinish + 1, Order_result_Column).Address).Value2
                realtime_checkbox = .Range(.Cells(RealtimeStart, realtime_checkbox_column).Address, .Cells(RealtimeFinish + 1, realtime_checkbox_column).Address).Value2
                doctrineinfo = .Range(.Cells(StockReadStart, Doctrine_Number_Column), .Cells(StockReadFinish + 1, Doctrine_Arg2_Column)).Value2
            End With
        Case LOAD_ARRAY_STOCKINFO
            With Sheets("Main")
                stockinfo = .Range(.Cells(StockReadStart, StockReadColumn).Address, .Cells(StockReadFinish + 1, TradeAmountColumn).Address).Value2
            End With
        Case LOAD_ARRAY_FAVORITE
            With Sheets("Main")
                favoriteinfo = .Range(.Cells(FavoriteStart, FavoriteReadColumn).Address, .Cells(FavoriteFinish + 1, Goodprice_deal_Favorite_Column).Address).Value2
            End With
        Case LOAD_ARRAY_BALANCE
            loadarray (LOAD_ARRAY_FAVORITE)
            loadarray (LOAD_ARRAY_AUTODEAL)
        Case LOAD_ARRAY_AUTODEAL
            loadarray (LOAD_ARRAY_STOCKINFO)
            loadarray (LOAD_ARRAY_DOCTRINES)
            With Sheets("Main")
                autodealinfo = .Range(.Cells(StockReadStart, Autodeal_CheckBox_Column).Address, .Cells(StockReadFinish + 1, Autodeal_status_col).Address).Value2
            End With
        Case LOAD_ARRAY_BUYSELL
            loadarray (LOAD_ARRAY_STOCKINFO)
            With Sheets("Main")
                buysellinfo = .Range(.Cells(StockReadStart, Order_Ckbox_Column).Address, .Cells(StockReadFinish + 1, Order_result_Column).Address).Value2
            End With
        Case LOAD_ARRAY_REALTIME
            With Sheets("Main")
                realtime_checkbox = .Range(.Cells(RealtimeStart, realtime_checkbox_column).Address, .Cells(RealtimeFinish + 1, realtime_checkbox_column).Address).Value2
            End With
        Case LOAD_ARRAY_DOCTRINES
            With Sheets("Main")
                doctrineinfo = .Range(.Cells(StockReadStart, Doctrine_Number_Column), .Cells(StockReadFinish + 1, Doctrine_Arg2_Column)).Value2
            End With
    End Select
    
End Sub

Sub Writearray(ByRef val As Integer)

    Application.ScreenUpdating = False
    
    Dim tmp() As Variant
    Dim c As Integer, r As Integer
    
    Select Case val
        Case LOAD_ARRAY_STOCKINFO
            With Sheets("Main")
                tmp = .Range(.Cells(StockReadStart, StockReadColumn + 1).Address, .Cells(StockReadFinish, TradeAmountColumn).Address).Value2
                For c = 1 To UBound(tmp, 2)
                        For r = 1 To UBound(tmp, 1)
                            tmp(r, c) = stockinfo(r, c + 1)
                        Next
                    Next
                .Range(.Cells(StockReadStart, StockReadColumn + 1).Address, .Cells(StockReadFinish, TradeAmountColumn).Address).FormulaR1C1 = tmp
            End With
        Case LOAD_ARRAY_FAVORITE
            With Sheets("Main")
                .Range(.Cells(FavoriteStart, FavoriteReadColumn).Address, .Cells(FavoriteFinish, Goodprice_deal_Favorite_Column).Address).FormulaR1C1 = favoriteinfo
            End With
        Case LOAD_ARRAY_BALANCE
            Writearray (LOAD_ARRAY_FAVORITE)
            With Sheets("Main")
                tmp = .Range(.Cells(StockReadStart, Autodeal_ownstock_col).Address, .Cells(StockReadFinish, Goodprice_deal_AUTO_Column).Address).Value2
                For c = 1 To UBound(tmp, 2)
                    For r = 1 To UBound(tmp, 1)
                        tmp(r, c) = autodealinfo(r, c + 1)
                    Next r
                Next c
                .Range(.Cells(StockReadStart, Autodeal_ownstock_col).Address, .Cells(StockReadFinish, Goodprice_deal_AUTO_Column).Address).FormulaR1C1 = tmp
            End With
        Case LOAD_ARRAY_AUTODEAL
            With Sheets("Main")
                tmp = .Range(.Cells(StockReadStart, Autodeal_status_col).Address, .Cells(StockReadFinish, Autodeal_status_col).Address).Value2
                For c = 1 To UBound(tmp, 2)
                    For r = 1 To UBound(tmp, 1)
                        tmp(r, c) = autodealinfo(r, iAutoStatusCol)
                    Next r
                Next c
                .Range(.Cells(StockReadStart, Autodeal_status_col).Address, .Cells(StockReadFinish, Autodeal_status_col).Address).Value2.FormulaR1C1 = tmp
            End With
        Case LOAD_ARRAY_BUYSELL
            With Sheets("Main")
                .Range(.Cells(StockReadStart, Order_Ckbox_Column).Address, .Cells(StockReadFinish, Order_result_Column).Address).FormulaR1C1 = buysellinfo
            End With
    End Select
    Application.ScreenUpdating = True

    End Sub

Sub Load_settings()
    
    'plugin use
    NOT_USE_PLUGIN = Not (GetSettings("USE_PLUGIN"))
    
    'favorite read
    FavoriteReadColumn = GetColumn("FavoriteStart")
    FavoriteStart = GetRow("FavoriteStart")
    FavoriteFinish = GetRow("FavoriteFinish")
    FavoriteWriteColumn = GetColumn("Favorite_target_column")
    FavoriteNameColumn = GetColumn("FavoriteNameColumn")
    
    'stocklist read
    StockReadColumn = GetColumn("StockreadStart")
    StockReadStart = GetRow("StockreadStart")
    StockReadFinish = GetRow("StockreadFinish")
    StockWriteColumn = GetColumn("Stockread_target_column")
    StockNameColumn = GetColumn("StockNameColumn")
   
    'real code read
    RealtimeReadColumn = GetColumn("RealtimeStart")
    RealtimeStart = GetRow("RealtimeStart")
    RealtimeFinish = GetRow("RealtimeFinish")
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
    
    loadarray LOAD_ARRAY_ALL
    
    Dim col_offset As Integer
    col_offset = TradeAmountColumn - UBound(stockinfo, 2)
    isReadCol = StockReadColumn - col_offset
    isNameCol = StockNameColumn - col_offset
    isWriteCol = StockWriteColumn - col_offset
    isAmountCol = TradeAmountColumn - col_offset

    col_offset = Goodprice_deal_Favorite_Column - UBound(favoriteinfo, 2)
    iFReadCol = FavoriteReadColumn - col_offset
    iFWriteCol = FavoriteWriteColumn - col_offset
    iFNameCol = FavoriteNameColumn - col_offset
    iFGoodPriceCol = Goodprice_deal_Favorite_Column - col_offset

    col_offset = Autodeal_status_col - UBound(autodealinfo, 2)
    iAutoBoxCol = Autodeal_CheckBox_Column - col_offset
    iAutoBadsellCol = Autodeal_badsell_Price_Col - col_offset
    iAutoGoodsellCol = Autodeal_goodsell_Price_Col - col_offset
    iAutoGoodPerCol = Autodeal_goodpercent_Col - col_offset
    iAutoBadPerCol = Autodeal_badpercent_Col - col_offset
    iAutoOwnCol = Autodeal_ownstock_col - col_offset
    iAutoStatusCol = Autodeal_status_col - col_offset
    iBalGoodPriceCol = Goodprice_deal_AUTO_Column - col_offset
    iBalOwnAmountCol = Autodeal_ownstock_col - col_offset
    
    col_offset = Order_result_Column - UBound(buysellinfo, 2)
    iOrderBoxCol = Order_Ckbox_Column - col_offset
    iOrderPriceCol = Order_price_Column - col_offset
    iOrderAmountCol = Order_amount_Column - col_offset
    iOrderMoneyCol = Order_money_Column - col_offset
    iOrderResultCol = Order_result_Column - col_offset
    
    col_offset = Doctrine_Arg2_Column - UBound(doctrineinfo, 2)
    iDocNumCol = Doctrine_Number_Column - col_offset
    iDocArg1Col = Doctrine_Arg1_Column - col_offset
    iDocArg2Col = Doctrine_Arg2_Column - col_offset
    
    
End Sub

Private Sub CheckBox4_Change()
    NOT_USE_PLUGIN = Not CheckBox4.Value
    Sheets("v").Range("USE_PLUGIN").Value2 = CheckBox4.Value
End Sub

Private Sub Autodeal1_Change()
    Autodeal_Change_Handler 0, Autodeal1.Value
End Sub

Private Sub Autodeal2_Change()
    Autodeal_Change_Handler 1, Autodeal2.Value
End Sub

Private Sub Autodeal3_Change()
    Autodeal_Change_Handler 2, Autodeal3.Value
End Sub

Private Sub Autodeal4_Change()
    Autodeal_Change_Handler 3, Autodeal4.Value
End Sub

Private Sub Autodeal5_Change()
    Autodeal_Change_Handler 4, Autodeal5.Value
End Sub

Private Sub Autodeal6_Change()
    Autodeal_Change_Handler 5, Autodeal6.Value
End Sub

Private Sub Autodeal7_Change()
    Autodeal_Change_Handler 6, Autodeal7.Value
End Sub

Private Sub Autodeal8_Change()
    Autodeal_Change_Handler 7, Autodeal8.Value
End Sub

Private Sub Autodeal9_Change()
    Autodeal_Change_Handler 8, Autodeal9.Value
End Sub

Private Sub Autodeal10_Change()
    Autodeal_Change_Handler 9, Autodeal10.Value
End Sub

Private Sub Autodeal11_Change()
    Autodeal_Change_Handler 10, Autodeal11.Value
End Sub

Private Sub Autodeal12_Change()
    Autodeal_Change_Handler 11, Autodeal12.Value
End Sub

Private Sub Autodeal13_Change()
    Autodeal_Change_Handler 12, Autodeal13.Value
End Sub

Private Sub Autodeal14_Change()
    Autodeal_Change_Handler 13, Autodeal14.Value
End Sub

Private Sub Autodeal15_Change()
    Autodeal_Change_Handler 14, Autodeal15.Value
End Sub

Private Sub Autodeal16_Change()
    Autodeal_Change_Handler 15, Autodeal16.Value
End Sub

Private Sub Autodeal17_Change()
    Autodeal_Change_Handler 16, Autodeal17.Value
End Sub

Private Sub Autodeal18_Change()
    Autodeal_Change_Handler 17, Autodeal18.Value
End Sub

Private Sub Autodeal19_Change()
    Autodeal_Change_Handler 18, Autodeal19.Value
End Sub

Private Sub Autodeal20_Change()
    Autodeal_Change_Handler 19, Autodeal20.Value
End Sub

Private Sub Realtime1_Change()
    Realtime_Change_Handler 0, Realtime1.Value
End Sub

Private Sub Realtime2_Change()
    Realtime_Change_Handler 1, Realtime2.Value
End Sub

Private Sub Realtime3_Change()
    Realtime_Change_Handler 2, Realtime3.Value
End Sub

Private Sub Realtime4_Change()
    Realtime_Change_Handler 3, Realtime4.Value
End Sub

Private Sub Realtime5_Change()
    Realtime_Change_Handler 4, Realtime5.Value
End Sub

Private Sub Realtime6_Change()
    Realtime_Change_Handler 5, Realtime6.Value
End Sub

Private Sub Realtime7_Change()
    Realtime_Change_Handler 6, Realtime7.Value
End Sub

Private Sub Realtime8_Change()
    Realtime_Change_Handler 7, Realtime8.Value
End Sub

Private Sub Realtime9_Change()
    Realtime_Change_Handler 8, Realtime9.Value
End Sub

Private Sub Realtime10_Change()
    Realtime_Change_Handler 9, Realtime10.Value
End Sub

Private Sub Realtime11_Change()
    Realtime_Change_Handler 10, Realtime11.Value
End Sub

Private Sub Realtime12_Change()
    Realtime_Change_Handler 11, Realtime12.Value
End Sub

Private Sub Realtime13_Change()
    Realtime_Change_Handler 12, Realtime13.Value
End Sub

Private Sub Realtime14_Change()
    Realtime_Change_Handler 13, Realtime14.Value
End Sub

Private Sub Realtime15_Change()
    Realtime_Change_Handler 14, Realtime15.Value
End Sub

Private Sub Realtime16_Change()
    Realtime_Change_Handler 15, Realtime16.Value
End Sub

Private Sub Realtime17_Change()
    Realtime_Change_Handler 16, Realtime17.Value
End Sub

Private Sub Realtime18_Change()
    Realtime_Change_Handler 17, Realtime18.Value
End Sub

Private Sub Realtime19_Change()
    Realtime_Change_Handler 18, Realtime19.Value
End Sub

Private Sub Realtime20_Change()
    Realtime_Change_Handler 19, Realtime20.Value
End Sub

Sub Realtime_Change_Handler(ByRef i As Integer, ByRef checked As Boolean)
    
    Load_settings
    Sheets("Main").Cells(i + RealtimeStart, realtime_checkbox_column).FormulaR1C1 = checked
    If CBool(GetSettings("realtime_enabled")) = False Then Exit Sub
    
    If UsePlugin() = True Then
        Call SendFunction("Realtime_" & CStr(checked) & "_" & CStr(i))
        Exit Sub
    End If
    
    With Sheets("Main")
        
        If checked = True Then
        
            'btnRequest_Click
            .Range(Cells(i + RealtimeStart, StockNameColumn).Address, Cells(i + RealtimeStart, TradeAmountColumn).Address).ClearContents
            loadarray (LOAD_ARRAY_STOCKINFO)
            
            Dim stockcode As String, col As Integer
            col = RealtimeReadColumn
            stockcode = .Cells(i + RealtimeStart, RealtimeReadColumn).Value2
            If Len(stockcode) = 6 And IsNumeric(stockcode) Then   'target.row, column
                GetStockName stockcode, i + RealtimeStart, StockNameColumn
                GetCurrentPrice stockcode, i + RealtimeStart, RealtimeWritecolumn
                GetRealtimePrice stockcode, i + RealtimeStart, RealtimeWritecolumn, i
            End If
            
            Application.OnTime Now() + TimeValue("0:00:05"), "VERIFY_NOT_ASSIGNED"

        Else
        
            StopRealtimePrice i
            
        End If
    End With

End Sub

Sub Autodeal_Change_Handler(ByRef i As Integer, ByRef checked As Boolean)

    Sheets("Main").Cells(i + RealtimeStart, GetColumnInMultiRange("SelectALL_AUTOdeal")).FormulaR1C1 = checked
    If UsePlugin = True Then Module1.SendFunction ("AutoRef_" & i)
    If CBool(GetSettings("Autodeal_Enabled")) = False Then Exit Sub
    
    ' do something
    
End Sub

'=========================
' 여기서부터 프로그램 구현하세요
'=========================
Private Sub CheckBox1_Click()

    Sheets("Main").Range(GetRange("SelectALL_deal")).FormulaR1C1 = CheckBox1.Value
        
End Sub

Private Sub CheckBox2_Click()

    'Sheets("Main").Range(GetRange("SelectALL_AUTOdeal")).FormulaR1C1 = CheckBox2.Value
    Application.ScreenUpdating = False
    
    Autodeal1.Value = CheckBox2.Value
    Autodeal2.Value = CheckBox2.Value
    Autodeal3.Value = CheckBox2.Value
    Autodeal4.Value = CheckBox2.Value
    Autodeal5.Value = CheckBox2.Value
    Autodeal6.Value = CheckBox2.Value
    Autodeal7.Value = CheckBox2.Value
    Autodeal8.Value = CheckBox2.Value
    Autodeal9.Value = CheckBox2.Value
    Autodeal10.Value = CheckBox2.Value
    Autodeal11.Value = CheckBox2.Value
    Autodeal12.Value = CheckBox2.Value
    Autodeal13.Value = CheckBox2.Value
    Autodeal14.Value = CheckBox2.Value
    Autodeal15.Value = CheckBox2.Value
    Autodeal16.Value = CheckBox2.Value
    Autodeal17.Value = CheckBox2.Value
    Autodeal18.Value = CheckBox2.Value
    Autodeal19.Value = CheckBox2.Value
    Autodeal20.Value = CheckBox2.Value
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CheckBox3_Click()
    
    'Sheets("Main").Range(GetRange("realtime_checkbox_range")).FormulaR1C1 = CheckBox3.Value
    
    Application.ScreenUpdating = False
    
    Realtime1.Value = CheckBox3.Value
    Realtime2.Value = CheckBox3.Value
    Realtime3.Value = CheckBox3.Value
    Realtime4.Value = CheckBox3.Value
    Realtime5.Value = CheckBox3.Value
    Realtime6.Value = CheckBox3.Value
    Realtime7.Value = CheckBox3.Value
    Realtime8.Value = CheckBox3.Value
    Realtime9.Value = CheckBox3.Value
    Realtime10.Value = CheckBox3.Value
    Realtime11.Value = CheckBox3.Value
    Realtime12.Value = CheckBox3.Value
    Realtime13.Value = CheckBox3.Value
    Realtime14.Value = CheckBox3.Value
    Realtime15.Value = CheckBox3.Value
    Realtime16.Value = CheckBox3.Value
    Realtime17.Value = CheckBox3.Value
    Realtime18.Value = CheckBox3.Value
    Realtime19.Value = CheckBox3.Value
    Realtime20.Value = CheckBox3.Value
    
    Application.ScreenUpdating = True
    
End Sub

Sub btnRequest_Click()
Attribute btnRequest_Click.VB_ProcData.VB_Invoke_Func = "q\n14"

    If UsePlugin Then
        Module1.SendFunction ("GetStockInfo")
        Exit Sub
    End If

    'Clear stock and Favorite
    With Sheets("Main")
        .Range(GetRange("Delete_Price_Stock")).ClearContents
        .Range(GetRange("Delete_Price_Favorite")).ClearContents
    End With


    Call loadarray(LOAD_ARRAY_STOCKINFO)
    Call loadarray(LOAD_ARRAY_FAVORITE)
    
    req_remaining(LOAD_ARRAY_STOCKINFO) = 0
    req_remaining(LOAD_ARRAY_FAVORITE) = 0
    
    Dim sCode As String, rows As Integer
    For rows = 1 To UBound(stockinfo, 1)
        stockinfo(rows, isNameCol) = vbNullString
        stockinfo(rows, isWriteCol) = vbNullString
        stockinfo(rows, isAmountCol) = vbNullString
        stockinfo(rows, iFNameCol) = vbNullString
        stockinfo(rows, iFWriteCol) = vbNullString
    Next rows

    For rows = 1 To UBound(stockinfo, 1)
        
        sCode = stockinfo(rows, isReadCol)
        If sCode <> vbNullString Or Len(sCode) = 6 Then
            'Get stock price and amount
            Call oShin_PriceAmount.SetQueryName("SC")
            Call oShin_PriceAmount.SetSingleData(0, CStr(sCode))
            stockinfo(rows, isWriteCol) = oShin_PriceAmount.RequestData()
            If stockinfo(rows, isWriteCol) <= 0 Then
                Send_ErrorHandler "현재가", CInt(stockinfo(rows, isWriteCol)), oShin_PriceAmount.GetErrorState() & Chr(10) & oShin_PriceAmount.GetErrorMessage(), "stockinfo,iswritecol"
                Exit Sub
            End If
            req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) + 1

            'and stock name
            Call oShin_name.SetQueryName("SB")
            Call oShin_name.SetSingleData(0, CStr(sCode))
            stockinfo(rows, isNameCol) = oShin_name.RequestData()
            If stockinfo(rows, isNameCol) <= 0 Then
                Send_ErrorHandler "현재가", CInt(stockinfo(rows, isNameCol)), oShin_name.GetErrorState() & Chr(10) & oShin_name.GetErrorMessage(), "stockinfo,isNameCol"
                Exit Sub
            End If
            req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) + 1

        End If

        sCode = favoriteinfo(rows, iFReadCol)
        If sCode <> vbNullString Or Len(sCode) = 6 Then
            'Get stock price (only)
            Call oShin_PriceOnly.SetQueryName("SC")
            Call oShin_PriceOnly.SetSingleData(0, CStr(sCode))
            favoriteinfo(rows, iFWriteCol) = oShin_PriceOnly.RequestData()
            If favoriteinfo(rows, iFWriteCol) <= 0 Then
                Send_ErrorHandler "현재가", CInt(stockinfo(rows, iFWriteCol)), oShin_PriceOnly.GetErrorState() & Chr(10) & oShin_PriceOnly.GetErrorMessage(), "stockinfo,iFWriteCol"
                Exit Sub
            End If
            req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) + 1

            'and stock name
            Call oShin_favname.SetQueryName("SB")
            Call oShin_favname.SetSingleData(0, CStr(sCode))
            favoriteinfo(rows, iFNameCol) = oShin_favname.RequestData()
            If favoriteinfo(rows, iFNameCol) <= 0 Then
                Send_ErrorHandler "현재가", CInt(stockinfo(rows, iFNameCol)), oShin_favname.GetErrorState() & Chr(10) & oShin_favname.GetErrorMessage(), "stockinfo,iFNameCol"
                Exit Sub
            End If
            req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) + 1

        End If
    
    Next rows

    timer_tick_count(LOAD_ARRAY_STOCKINFO) = 0
    Application.OnTime Sheets("v").Range("crazybuy_next_ontime").Value2, "Timer_StockInfo_Tick"
    'Timer_StockInfo.Interval = 100
    'Timer_StockInfo.Enabled = True
    


    Application.OnTime Now() + TimeValue("0:00:05"), "VERIFY_NOT_ASSIGNED"
    'Exit Sub
    
    'Load stock first
    'Dim rge As Range, stockcode As String
    'For Each rge In Sheets("Main").Range(Sheets("Main").Cells(StockReadStart, StockReadColumn).Address, Sheets("Main").Cells(StockReadFinish, StockReadColumn).Address)
    '    With rge
    '        stockcode = .Value2
    '        If Len(stockcode) = 6 And IsNumeric(stockcode) Then  'row, column
    '            GetStockName stockcode, .row, StockNameColumn
    '            GetCurrentPrice stockcode, .row, StockWriteColumn
    '            GetTradeAmount stockcode, .row, TradeAmountColumn
    '        End If
    '    End With
    'Next rge
    
    'load favorite
    'For Each rge In Sheets("Main").Range(Sheets("Main").Cells(FavoriteStart, FavoriteReadColumn).Address, Sheets("Main").Cells(FavoriteFinish, FavoriteReadColumn).Address)
    '    With rge
    '        stockcode = .Value2
    '        If Len(stockcode) = 6 And IsNumeric(stockcode) Then  'row, column
    '            GetStockName stockcode, .row, FavoriteNameColumn
    '            GetCurrentPrice stockcode, .row, FavoriteWriteColumn
    '        End If
    '    End With
    'Next rge
    'Application.OnTime Now() + TimeValue("0:00:05"), "VERIFY_NOT_ASSIGNED"
    
End Sub

Sub btnRequest_realtime_1_Click()
Attribute btnRequest_realtime_1_Click.VB_ProcData.VB_Invoke_Func = "R\n14"

    If UsePlugin() = True Then
        SetSettings "realtime_enabled", True, False
        Call SendFunction("Realtime_StartAll")
        Exit Sub
    End If

    btnRequest_Click
    'Sheets("Main").Range(GetRange("Delete_Price_Realtime")).ClearContents
    
    Dim rge As Range, stockcode As String
    For Each rge In Sheets("Main").Range(Sheets("Main").Cells(RealtimeStart, RealtimeReadColumn).Address, Sheets("Main").Cells(RealtimeFinish, RealtimeReadColumn).Address)
        With rge
            stockcode = .Value2
            If Sheets("Main").Cells(.row, realtime_checkbox_column).Value2 = True And Len(stockcode) = 6 And IsNumeric(stockcode) Then   'row, column
                'GetStockName stockcode, .row, StockNameColumn
                'GetCurrentPrice stockcode, .row, RealtimeWritecolumn
                GetRealtimePrice stockcode, .row, RealtimeWritecolumn, .row - RealtimeStart
            End If
        End With
    Next rge
    
    If Sheets("v").Range("realtime_enabled").Value2 = False Then Application.OnTime Sheets("v").Range("realtime_next_runtime").Value2, "Realtime_Update"
    SetSettings "realtime_enabled", True, False
    Application.OnTime Now() + TimeValue("0:00:05"), "VERIFY_NOT_ASSIGNED"
    
End Sub

Sub btnstop_realtime_1_click()
Attribute btnstop_realtime_1_click.VB_ProcData.VB_Invoke_Func = "T\n14"

    If UsePlugin() = True Then
        SetSettings "realtime_enabled", False, False
        Call SendFunction("Realtime_StopAll")
        Exit Sub
    End If
    
    Dim i As Integer
    SetSettings "realtime_enabled", False, False
    For i = 0 To 19
        StopRealtimePrice i
    Next i
    
End Sub

Sub button_buysell_bacckend(ByRef buysell As Integer)
        
    Dim row As Integer, col As Integer, rge As Range
    Dim price As String
    
    col = Order_Ckbox_Column

    With Sheets("Main")
        
        For Each rge In .Range(Sheets("Main").Cells(StockReadStart, col).Address, Sheets("Main").Cells(StockReadFinish, col).Address)
            If rge.Value2 = True Then
                '시장가 혹은 지정가 계산
                row = rge.row
                price = .Cells(row, Order_price_Column).Value2
                If price = vbNullString Then
                    MsgBox "구매가격을 넣어라"
                    Exit Sub
                End If
                If price = "시장가" Or price = "ㅅ" Or price = "시" Then price = PRICE_MARKET_PRICE
                If IsNumeric(price) And CLng(price) <> 0 And CLng(price) >= PRICE_MARKET_PRICE Then
                
                    '둘다 안넣은 경우
                    If .Cells(row, Order_money_Column).Value2 = vbNullString And .Cells(row, Order_amount_Column).Value2 = vbNullString Then
                        MsgBox "둘중에 하나는 넣어야지 ㅄ"
                        Exit Sub
                    End If
                    
                    '시장가 아닌경우
                    If price <> PRICE_MARKET_PRICE Then
                    
                        '주만 넣은 경우 금액을 계산
                        If CLng(.Cells(row, Order_money_Column).Value2) = 0 And CLng(.Cells(row, Order_amount_Column).Value2) > 0 Then
                            .Cells(row, Order_money_Column).Value2 = CLng(.Cells(row, Order_amount_Column).Value2 * price)
                        '금액만 넣은 경우 주를 계산
                        ElseIf CLng(.Cells(row, Order_money_Column).Value2) > 0 And CLng(.Cells(row, Order_amount_Column).Value2) = 0 Then
                            .Cells(row, Order_amount_Column).Value2 = CLng(Application.WorksheetFunction.RoundDown(CDbl(.Cells(row, Order_money_Column).Value2) / CDbl(price), 0))
                        End If
                    Else
                        '시장가인데 현재가가 쓰여 있고 금액만 넣은 경우 금액 / 현재가를 나눈 주를집어넣
                        'MsgBox .Cells(row, StockWriteColumn).Value2
                        If CLng(.Cells(row, StockWriteColumn).Value2) > 0 And CLng(.Cells(row, Order_money_Column).Value2) > 0 And CLng(.Cells(row, Order_amount_Column).Value2) = 0 Then
                            .Cells(row, Order_amount_Column).Value2 = CLng(Application.WorksheetFunction.RoundDown(CDbl(.Cells(row, Order_money_Column).Value2) / CDbl(.Cells(row, StockWriteColumn).Value2), 0))
                        End If
                        
                    End If
    
                    '계산된 가격과 주로 주문 넣기
                    Select Case buysell
                        Case ORDER_BUY: BuyStock .Cells(row, StockReadColumn).Value2, CLng(price), .Cells(row, Order_amount_Column).Value2, row, Order_result_Column, False
                        Case ORDER_SELL: SellStock .Cells(row, StockReadColumn).Value2, CLng(price), .Cells(row, Order_amount_Column).Value2, row, Order_result_Column, False
                    End Select
                
                    '주문넣었으면 제거
                    .Cells(row, col).FormulaR1C1 = False
                End If
            End If
        Next
        

        
    End With
    
    
End Sub

Sub button_sellstock_click()
Attribute button_sellstock_click.VB_ProcData.VB_Invoke_Func = "E\n14"
    
    If UsePlugin = True Then
        Module1.SendFunction ("SellStock")
        Exit Sub
    End If
    
    button_buysell_bacckend ORDER_SELL
    
End Sub

Sub button_buystock_click()
Attribute button_buystock_click.VB_ProcData.VB_Invoke_Func = "W\n14"
    
    If UsePlugin = True Then
        Module1.SendFunction ("BuyStock")
        Exit Sub
    End If
    
    button_buysell_bacckend ORDER_BUY
    
End Sub

Sub enable_auto_trading()
Attribute enable_auto_trading.VB_ProcData.VB_Invoke_Func = "A\n14"
    
    '손절가 / 익절가 가계산 실시
    Dim i As Integer, goodprice
    
    With Sheets("Main")
        For i = RealtimeStart To RealtimeFinish
        
            '작동여부 관계없이 제비용단가 적힌 보유주식은 모두 계산해둠
            goodprice = .Cells(i, Goodprice_deal_AUTO_Column).Value2
            If goodprice <> vbNullString And IsNumeric(goodprice) And goodprice > 0 Then
                
                '익절가 퍼템만 있고 값이 없는 경우 익절가 자동계산
                If .Cells(i, Autodeal_goodsell_Price_Col).Value2 = vbNullString And .Cells(i, Autodeal_goodpercent_Col).Value2 <> vbNullString Then
                    .Cells(i, Autodeal_goodsell_Price_Col).Value2 = goodprice + Round(CDbl(goodprice) * (CDbl(.Cells(i, Autodeal_goodpercent_Col)) / 100), 0)
                End If
    
                '익절값만 있고 퍼템이 없는경우 익절퍼 자동계산
                If .Cells(i, Autodeal_goodsell_Price_Col).Value2 <> vbNullString And .Cells(i, Autodeal_goodpercent_Col).Value2 = vbNullString Then
                    .Cells(i, Autodeal_goodpercent_Col).Value2 = Round(((CDbl(.Cells(i, Autodeal_goodsell_Price_Col).Value2) / CDbl(goodprice)) - 1) * 100, 1)
                End If
                
                '손절가 퍼템만 있고 값이 없는 경우 익절가 자동계산
                If .Cells(i, Autodeal_badsell_Price_Col).Value2 = vbNullString And .Cells(i, Autodeal_badpercent_Col).Value2 <> vbNullString Then
                    .Cells(i, Autodeal_badsell_Price_Col).Value2 = goodprice + Round(CDbl(goodprice) * (CDbl(.Cells(i, Autodeal_badpercent_Col)) / 100), 0)
                End If
    
                '손절값만 있고 퍼템이 없는경우 익절퍼 자동계산
                If .Cells(i, Autodeal_badsell_Price_Col).Value2 <> vbNullString And .Cells(i, Autodeal_badpercent_Col).Value2 = vbNullString Then
                    .Cells(i, Autodeal_badpercent_Col).Value2 = Round(((CDbl(.Cells(i, Autodeal_badsell_Price_Col).Value2) / CDbl(goodprice)) - 1) * 100, 1)
                End If
                
            End If
        Next i
    End With
    
    'check if realtime is working
    Sheets("v").Range("Autodeal_Enabled").FormulaR1C1 = True
    If UsePlugin = True Then Module1.SendFunction ("AutoDeal_Start")
    
End Sub

Sub disable_auto_trading()
Attribute disable_auto_trading.VB_ProcData.VB_Invoke_Func = "S\n14"
        
    Sheets("v").Range("Autodeal_Enabled").FormulaR1C1 = False
    If UsePlugin = True Then Module1.SendFunction ("AutoDeal_STOP")
    
End Sub

Sub enable_analysis()
Attribute enable_analysis.VB_ProcData.VB_Invoke_Func = "V\n14"
    
    Sheets("v").Range("Analysis_Enabled").FormulaR1C1 = True
    
End Sub

Sub disable_analysis()
Attribute disable_analysis.VB_ProcData.VB_Invoke_Func = "B\n14"

    Sheets("v").Range("Analysis_Enabled").FormulaR1C1 = False

End Sub

Sub GetCurrentBalance()
Attribute GetCurrentBalance.VB_ProcData.VB_Invoke_Func = "Q\n14"

    If UsePlugin Then
        Module1.SendFunction ("GetBalanceInfo")
        Exit Sub
    End If

    Dim i As Integer, ret As Integer
    i = Sheets("v").Range("Selected_Account").Value2
    
    Call objShinhan_balance.SetQueryName("SABA609Q1")
    Call objShinhan_balance.SetSingleData(0, CStr(Sheets("v").Range("Account_number_" & i).Value2))
    Call objShinhan_balance.SetSingleData(1, "redacted")
    Call objShinhan_balance.SetSingleData(2, "redacted")
    Call objShinhan_balance.SetSingleData(3, "redacted")
    Call objShinhan_balance.SetSingleData(4, "redacted")
    Call objShinhan_balance.SetSingleData(5, "redacted")
    Call objShinhan_balance.SetSingleData(6, "redacted")
    ret = objShinhan_balance.RequestData
    
    If ret <= 0 Then Send_ErrorHandler "잔고조회", objShinhan_balance.GetErrorState(), objShinhan_balance.GetErrorMessage(), CStr(Sheets("v").Range("Account_number_" & i).Value2)
            
End Sub

Sub crazybuy_onmymark_click()

    With Sheets("v")
        If .Range("crazybuy_running").Value2 = True Then
            MsgBox "이미 단주매매가 실행중입니다.", vbOKOnly + vbExclamation, "실행 중"
            Exit Sub
        End If
            
        Dim code As String
        code = .Range("crazybuy_code").Value2
        If Len(Trim(code)) <> 6 Then
            MsgBox "[단주매매] 종목코드가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
        
        Dim buyprice As String
        buyprice = .Range("crazybuy_buy_price").Value2
        If buyprice = "시장가" Then buyprice = CStr(PRICE_MARKET_PRICE)
        
        If Not IsNumeric(buyprice) Or CLng(buyprice) = 0 Or CLng(buyprice) < PRICE_MARKET_PRICE Then
            MsgBox "[단주매매] 구매가격이 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
        
        Dim sellprice As String
        sellprice = .Range("crazybuy_sell_price").Value2
        If sellprice = "시장가" Then sellprice = CStr(PRICE_MARKET_PRICE)
    End With
    
    If Not IsNumeric(sellprice) Or CLng(sellprice) = 0 Or CLng(sellprice) < PRICE_MARKET_PRICE Then
        MsgBox "[단주매매] 판매가격이 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If

    Module1.crazybuy_assign CRAZYBUY_TIME_BUY, code, CLng(buyprice), CLng(sellprice)
    MsgBox "철컥" & Chr(10) & "Get READY FOR THE NEXT BATTLE"
    SetSettings "crazybuy_iteration_time", 0, True

End Sub

Sub crazybuy_WARP_click()
Attribute crazybuy_WARP_click.VB_ProcData.VB_Invoke_Func = "C\n14"

    Module1.crazybuy_onetime

End Sub

Sub crazybuy_start_click()
Attribute crazybuy_start_click.VB_ProcData.VB_Invoke_Func = "Z\n14"
    
    With Sheets("v")
        If .Range("crazybuy_running").Value2 = True Then
            MsgBox "이미 단주매매가 실행중입니다.", vbOKOnly + vbExclamation, "실행 중"
            Exit Sub
        End If
            
        Dim code As String
        code = .Range("crazybuy_code").Value2
        If Len(Trim(code)) <> 6 Then
            MsgBox "[단주매매] 종목코드가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
        
        Dim buyprice As String
        buyprice = .Range("crazybuy_buy_price").Value2
        If buyprice = "시장가" Then buyprice = CStr(PRICE_MARKET_PRICE)
        
        If Not IsNumeric(buyprice) Or CLng(buyprice) = 0 Or CLng(buyprice) < PRICE_MARKET_PRICE Then
            MsgBox "[단주매매] 구매가격이 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
        
        Dim sellprice As String
        sellprice = .Range("crazybuy_sell_price").Value2
        If sellprice = "시장가" Then sellprice = CStr(PRICE_MARKET_PRICE)
        
        If Not IsNumeric(sellprice) Or CLng(sellprice) = 0 Or CLng(sellprice) < PRICE_MARKET_PRICE Then
            MsgBox "[단주매매] 판매가격이 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
        
        If CLng(.Range("crazybuy_iteration_time").Value2) < 1 Or CLng(.Range("crazybuy_iteration_time").Value2) Mod 2 = 1 Then
            MsgBox "[단주매매] 반복횟수가 올바르지 않습니다. 짝수로 입력하세요.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
    
        Module1.crazybuy_assign CRAZYBUY_TIME_BUY, code, CLng(buyprice), CLng(sellprice)
        .Range("crazybuy_running").FormulaR1C1 = True
        Application.OnTime .Range("crazybuy_running").Value2, "Crazybuy_TICK"
    End With
    
    'test1
    'Dim i As Integer
    'Dim CRAZYBUY_thistime As Boolean
    'CRAZYBUY_thistime = CRAZYBUY_TIME_BUY
    
    'For i = 0 To Sheets("v").Range("crazybuy_iteration_time").Value2
    
        'Select Case CRAZYBUY_thistime
            
            'Case CRAZYBUY_TIME_BUY:          BuyStock code, CLng(buyprice), 1, 1, 1, NOT_AUTOSTOP
            'Case CRAZYBUY_TIME_SELL:        SellStock code, CLng(sellprice), 1, 1, 1, NOT_AUTOSTOP
            
        'End Select
        
        'CRAZYBUY_thistime = Not CRAZYBUY_thistime
        'Sheets("v").Range("crazybuy_iteration_time").Value2 = CLng(Sheets("v").Range("crazybuy_iteration_time").Value2) - 1
        'Sheets("v").Range("crazybuy_next_ontime").Calculate
        'Sleep (1000)
        'Application.Wait (Sheets("v").Range("crazybuy_next_ontime").Value2)
        
    'Next i

    
End Sub

Sub crazybuy_stop_click()
Attribute crazybuy_stop_click.VB_ProcData.VB_Invoke_Func = "X\n14"
    
    Sheets("v").Range("crazybuy_running").FormulaR1C1 = False
    
End Sub




' ==========================================
' Background FUNTIONS
' ==========================================
Sub StopALL()
Attribute StopALL.VB_ProcData.VB_Invoke_Func = "D\n14"
    
    On Error Resume Next
    Load_settings
    
    objShinhan.UnRequestRTRegAll
    objShinhan_balance.UnRequestRTRegAll
    objShinhan_RT.UnRequestRTRegAll
    objShinhan_RT2.UnRequestRTRegAll
    objShinhan_RT3.UnRequestRTRegAll
    objShinhan_RT4.UnRequestRTRegAll
    objShinhan_RT5.UnRequestRTRegAll
    objShinhan_RT6.UnRequestRTRegAll
    objShinhan_RT7.UnRequestRTRegAll
    objShinhan_RT8.UnRequestRTRegAll
    objShinhan_RT9.UnRequestRTRegAll
    objShinhan_RT10.UnRequestRTRegAll
    objShinhan_RT11.UnRequestRTRegAll
    objShinhan_RT12.UnRequestRTRegAll
    objShinhan_RT13.UnRequestRTRegAll
    objShinhan_RT14.UnRequestRTRegAll
    objShinhan_RT15.UnRequestRTRegAll
    objShinhan_RT16.UnRequestRTRegAll
    objShinhan_RT17.UnRequestRTRegAll
    objShinhan_RT18.UnRequestRTRegAll
    objShinhan_RT19.UnRequestRTRegAll
    objShinhan_RT20.UnRequestRTRegAll
    
    MakeControl
    
    'objShinhan.SelfMemFree (True)
    'objShinhan_RT2.SelfMemFree (True)
    'objShinhan_RT3.SelfMemFree (True)
    'objShinhan_RT4.SelfMemFree (True)
    'objShinhan_RT5.SelfMemFree (True)
    'objShinhan_RT6.SelfMemFree (True)
    'objShinhan_RT7.SelfMemFree (True)
    'objShinhan_RT8.SelfMemFree (True)
    'objShinhan_RT9.SelfMemFree (True)
    'objShinhan_RT10.SelfMemFree (True)
    
    Dim i As Integer
    
    For i = 0 To RQRTSIZE
        RQ_CURPRICE_REALTIME(i, REQ_ID) = False
    Next i
    
    For i = 0 To RQSIZE
        RQ_ONCE_VARIABLE(i, REQ_ID) = REQ_INVALID_ID
    Next i
    
    MsgBox "모든 실시간 요청을 해제하였습니다.", vbOKOnly + vbInformation, "Complete"
    
End Sub

Sub StopRealtimePrice(ByRef realcode As Integer)

    If realcode < 0 Or realcode >= RQRTSIZE Then
        MsgBox "올바르지 않은 실시간 번호입니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If

    If RQ_CURPRICE_REALTIME(realcode, REQ_ID) = False Then Exit Sub
    
    Dim response As Boolean
    Select Case realcode
        Case 0: response = objShinhan_RT.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 1: response = objShinhan_RT2.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 2: response = objShinhan_RT3.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 3: response = objShinhan_RT4.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 4: response = objShinhan_RT5.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 5: response = objShinhan_RT6.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 6: response = objShinhan_RT7.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 7: response = objShinhan_RT8.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 8: response = objShinhan_RT9.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 9: response = objShinhan_RT10.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 10: response = objShinhan_RT11.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 11: response = objShinhan_RT12.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 12: response = objShinhan_RT13.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 13: response = objShinhan_RT14.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 14: response = objShinhan_RT15.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 15: response = objShinhan_RT16.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 16: response = objShinhan_RT17.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 17: response = objShinhan_RT18.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 18: response = objShinhan_RT19.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
        Case 19: response = objShinhan_RT20.UnRequestRTReg("SC", CStr(RQ_CURPRICE_REALTIME(realcode, REQ_STOCKCODE)))
    End Select
    
    If response = False Then
        MsgBox "종목코드 중단에 실패하였습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    RQ_CURPRICE_REALTIME(realcode, REQ_ID) = False
    
End Sub

Sub GetRealtimePrice(ByRef iCode As String, ByRef row As Integer, ByRef col As Integer, ByRef i As Integer)
    
    If Len(Trim(iCode)) <> 6 Then
        MsgBox "[실시간] " & i & "번 종목코드가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If row < 1 Or col < 1 Then
        MsgBox "[실시간] " & i & "번 업데이트 대상위치가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If i < 0 Or i >= RQRTSIZE Then
        MsgBox "[실시간] " & i & "번 올바르지 않은 실시간 요청코드입니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If RQ_CURPRICE_REALTIME(i, REQ_ID) = True Then
        MsgBox "해당 실시간 조회는 이미 작동중입니다." & Chr(10) & "먼저 종료해 주세요.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If

    Select Case i
        Case 0: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT.RequestRTReg("SC", iCode)
        Case 1: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT2.RequestRTReg("SC", iCode)
        Case 2: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT3.RequestRTReg("SC", iCode)
        Case 3: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT4.RequestRTReg("SC", iCode)
        Case 4: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT5.RequestRTReg("SC", iCode)
        Case 5: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT6.RequestRTReg("SC", iCode)
        Case 6: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT7.RequestRTReg("SC", iCode)
        Case 7: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT8.RequestRTReg("SC", iCode)
        Case 8: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT9.RequestRTReg("SC", iCode)
        Case 9: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT10.RequestRTReg("SC", iCode)
        Case 10: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT11.RequestRTReg("SC", iCode)
        Case 11: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT12.RequestRTReg("SC", iCode)
        Case 12: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT13.RequestRTReg("SC", iCode)
        Case 13: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT14.RequestRTReg("SC", iCode)
        Case 14: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT15.RequestRTReg("SC", iCode)
        Case 15: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT16.RequestRTReg("SC", iCode)
        Case 16: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT17.RequestRTReg("SC", iCode)
        Case 17: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT18.RequestRTReg("SC", iCode)
        Case 18: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT19.RequestRTReg("SC", iCode)
        Case 19: RQ_CURPRICE_REALTIME(i, REQ_ID) = objShinhan_RT20.RequestRTReg("SC", iCode)
    End Select
        
    If RQ_CURPRICE_REALTIME(i, REQ_ID) = False Then
        Dim nErr As Integer
        Dim sErrMsg As String
        
        Select Case i
            Case 0
                nErr = objShinhan_RT.GetErrorState()
                sErrMsg = objShinhan_RT.GetErrorMessage()
            Case 1
                nErr = objShinhan_RT2.GetErrorState()
                sErrMsg = objShinhan_RT2.GetErrorMessage()
            Case 2
                nErr = objShinhan_RT3.GetErrorState()
                sErrMsg = objShinhan_RT3.GetErrorMessage()
            Case 3
                nErr = objShinhan_RT4.GetErrorState()
                sErrMsg = objShinhan_RT4.GetErrorMessage()
            Case 4
                nErr = objShinhan_RT5.GetErrorState()
                sErrMsg = objShinhan_RT5.GetErrorMessage()
            Case 5
                nErr = objShinhan_RT6.GetErrorState()
                sErrMsg = objShinhan_RT6.GetErrorMessage()
            Case 6
                nErr = objShinhan_RT7.GetErrorState()
                sErrMsg = objShinhan_RT7.GetErrorMessage()
            Case 7
                nErr = objShinhan_RT8.GetErrorState()
                sErrMsg = objShinhan_RT8.GetErrorMessage()
            Case 8
                nErr = objShinhan_RT9.GetErrorState()
                sErrMsg = objShinhan_RT9.GetErrorMessage()
            Case 9
                nErr = objShinhan_RT10.GetErrorState()
                sErrMsg = objShinhan_RT10.GetErrorMessage()
            Case 10
                nErr = objShinhan_RT11.GetErrorState()
                sErrMsg = objShinhan_RT11.GetErrorMessage()
            Case 11
                nErr = objShinhan_RT12.GetErrorState()
                sErrMsg = objShinhan_RT12.GetErrorMessage()
            Case 12
                nErr = objShinhan_RT13.GetErrorState()
                sErrMsg = objShinhan_RT13.GetErrorMessage()
            Case 13
                nErr = objShinhan_RT14.GetErrorState()
                sErrMsg = objShinhan_RT14.GetErrorMessage()
            Case 14
                nErr = objShinhan_RT15.GetErrorState()
                sErrMsg = objShinhan_RT15.GetErrorMessage()
            Case 15
                nErr = objShinhan_RT16.GetErrorState()
                sErrMsg = objShinhan_RT16.GetErrorMessage()
            Case 16
                nErr = objShinhan_RT17.GetErrorState()
                sErrMsg = objShinhan_RT17.GetErrorMessage()
            Case 17
                nErr = objShinhan_RT18.GetErrorState()
                sErrMsg = objShinhan_RT18.GetErrorMessage()
            Case 18
                nErr = objShinhan_RT19.GetErrorState()
                sErrMsg = objShinhan_RT19.GetErrorMessage()
            Case 19
                nErr = objShinhan_RT20.GetErrorState()
                sErrMsg = objShinhan_RT20.GetErrorMessage()
        End Select
        
        Send_ErrorHandler "실시간조회", nErr, sErrMsg, CStr(i)
        Exit Sub
    End If
    
    RQ_CURPRICE_REALTIME(i, REQ_WRITE_ROW) = row
    RQ_CURPRICE_REALTIME(i, REQ_WRITE_COL) = col
    RQ_CURPRICE_REALTIME(i, REQ_STOCKCODE) = CStr(iCode)
    
End Sub

Sub Send_ErrorHandler(ByRef header As String, ByRef nErr As Integer, ByRef sErrMsg As String, ByRef additional As String)

    Select Case nErr
            Case 1: MsgBox "신한아이 인디를 실행하십시오.", vbOKOnly + vbExclamation, "Error"
            Case Else: MsgBox "[" & header & "] 송신오류 발생" & Chr(10) & nErr & Chr(10) & sErrMsg & Chr(10) & "보조자료 : " & additional, vbOKOnly + vbExclamation, "Error"
    End Select
    
End Sub

Sub Receive_ErrorHandler(ByRef header As String, ByRef nErr As Integer, ByRef sErrMsg As String, ByRef additional As String)
    
    MsgBox "[" & header & "] 수신오류 발생 : " & nErr & Chr(10) & sErrMsg & Chr(10) & "보조자료 : " & additional, vbOKOnly + vbExclamation, "Error"

End Sub


Sub BuyStock(ByRef iCode As String, ByRef price As Long, ByRef amount As Long, ByRef row As Integer, ByRef col As Integer, ByRef autostop As Integer)
    
    If Len(Trim(iCode)) <> 6 Then
        MsgBox "[매수] 종목코드가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If price = 0 Or price < PRICE_MARKET_PRICE Then
        MsgBox "[매수] 가격이 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If amount < 1 Then
        MsgBox "[매수] 판매주수가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If row < 1 Or col < 1 Then
        MsgBox "[매수] 업데이트 대상위치가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    BuySell_Backend ORDER_BUY, iCode, price, amount, row, col, autostop

End Sub

Sub SellStock(ByRef iCode As String, ByRef price As Long, ByRef amount As Long, ByRef row As Integer, ByRef col As Integer, ByRef autostop As Integer)
    
    If Len(Trim(iCode)) <> 6 Then
        MsgBox "[매도] 종목코드가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If price = 0 Or price < PRICE_MARKET_PRICE Then
        MsgBox "[매도] 가격이 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If amount < 1 Then
        MsgBox "[매도] 판매주수가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If row < 1 Or col < 1 Then
        MsgBox "[매도] 업데이트 대상위치가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    BuySell_Backend ORDER_SELL, iCode, price, amount, row, col, autostop

End Sub

Sub BuySell_Backend(ByRef buyorsell As Integer, ByRef iCode As String, ByRef price As Long, ByRef amount As Long, ByRef row As Integer, ByRef col As Integer, ByRef autostop As Integer)

    Dim i As Integer
        For i = 0 To RQSIZE
            If RQ_BUYSELL_VARIABLE(i, REQ_ID) <= REQ_INVALID_ID Then Exit For
        Next i
        If i = RQSIZE Then
            MsgBox "[구매/판매] 현재 요청을 더 할 수 없습니다." & Chr(10) & "잠시후 다시 시도해주세요.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
    End If

    Call objShinhan_buysell.SetQueryName("SABA101U1")
    Call objShinhan_buysell.SetSingleData(0, CStr(Sheets("v").Range("Account_number_" & Sheets("v").Range("Selected_Account").Value2).Value2))
    Call objShinhan_buysell.SetSingleData(1, "redacted")
    Call objShinhan_buysell.SetSingleData(2, "redacted")
    Call objShinhan_buysell.SetSingleData(3, vbNullString)
    Call objShinhan_buysell.SetSingleData(4, vbNullString)
    Call objShinhan_buysell.SetSingleData(5, "0")
    Call objShinhan_buysell.SetSingleData(6, "00")
    Call objShinhan_buysell.SetSingleData(7, CStr(buyorsell))
    Call objShinhan_buysell.SetSingleData(8, "A" & iCode)
    Call objShinhan_buysell.SetSingleData(9, CStr(amount))
    If price = PRICE_MARKET_PRICE Then Call objShinhan_buysell.SetSingleData(10, "0") Else Call objShinhan_buysell.SetSingleData(10, CStr(price))
    Call objShinhan_buysell.SetSingleData(11, "1")
    If price = PRICE_MARKET_PRICE Then Call objShinhan_buysell.SetSingleData(12, "1") Else Call objShinhan_buysell.SetSingleData(12, "2")
    Call objShinhan_buysell.SetSingleData(13, "0")
    Call objShinhan_buysell.SetSingleData(14, "0")
    Call objShinhan_buysell.SetSingleData(15, vbNullString)
    Call objShinhan_buysell.SetSingleData(16, vbNullString)
    Call objShinhan_buysell.SetSingleData(21, "Y")

    '매수매도
    RQ_BUYSELL_VARIABLE(i, REQ_ID) = objShinhan_buysell.RequestData
    If RQ_BUYSELL_VARIABLE(i, REQ_ID) <= 0 Then
        Send_ErrorHandler "구매/판매", objShinhan_buysell.GetErrorState(), objShinhan_buysell.GetErrorMessage(), _
            "buysell = " & buyorsell & "종목코드 = " & iCode & ", 가격 = " & price & ", 주 = " & amount & ", row = " & row & ", col = " & col
        Exit Sub
    End If
    
    Dim msg As String
    Select Case autostop
        Case NOT_AUTOSTOP: If buyorsell = ORDER_BUY Then msg = "매수요청" Else msg = "매도요청"
        Case AUTOSTOP_GOOD: msg = "익절요청"
        Case AUTOSTOP_BAD: msg = "손절요청"
    End Select
    Sheets("Main").Cells(row, col).FormulaR1C1 = msg
    
    RQ_BUYSELL_VARIABLE(i, REQ_WRITE_ROW) = row
    RQ_BUYSELL_VARIABLE(i, REQ_WRITE_COL) = col
    RQ_BUYSELL_VARIABLE(i, REQ_TYPE) = buyorsell
    RQ_BUYSELL_VARIABLE(i, REQ_STOCKCODE) = CStr(iCode)
    RQ_BUYSELL_VARIABLE(i, REQ_PRICE) = price
    RQ_BUYSELL_VARIABLE(i, REQ_AMOUNT) = amount
    RQ_BUYSELL_VARIABLE(i, REQ_AUTOSTOP) = autostop
        
End Sub

Sub GetStockName(ByRef iCode As String, ByRef row As Integer, ByRef col As Integer)
    Request_backend iCode, row, col, REQ_TYPE_NAME
End Sub

Sub GetCurrentPrice(ByRef iCode As String, ByRef row As Integer, ByRef col As Integer)
    Request_backend iCode, row, col, REQ_TYPE_PRICE
End Sub

Sub GetTradeAmount(ByRef iCode As String, ByRef row As Integer, ByRef col As Integer)
    Request_backend iCode, row, col, REQ_TYPE_TRADEAMOUNT
End Sub

Private Sub Request_backend(ByRef iCode As String, ByRef row As Integer, ByRef col As Integer, ByRef reqtype As Integer)
    
    If Len(Trim(iCode)) <> 6 Then
        MsgBox "[현재가] 종목코드가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If row < 1 Or col < 1 Then
        MsgBox "[현재가] 업데이트 대상위치가 올바르지 않습니다.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
        
    Dim i As Integer
    For i = 0 To RQSIZE
        If RQ_ONCE_VARIABLE(i, REQ_ID) <= REQ_INVALID_ID Then Exit For
    Next i
    If i = RQSIZE Then
        MsgBox "현재 요청을 더 할 수 없습니다." & Chr(10) & "잠시후 다시 시도해주세요.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
        
    Select Case reqtype
        Case REQ_TYPE_NAME: Call objShinhan.SetQueryName("SB")
        Case REQ_TYPE_PRICE: Call objShinhan.SetQueryName("SC")
        Case REQ_TYPE_TRADEAMOUNT: Call objShinhan.SetQueryName("SC")
    End Select
    
    Call objShinhan.SetSingleData(0, iCode)
    
    RQ_ONCE_VARIABLE(i, REQ_ID) = objShinhan.RequestData
    If RQ_ONCE_VARIABLE(i, REQ_ID) <= 0 Then
        Send_ErrorHandler "이름/현재가", objShinhan.GetErrorState(), objShinhan.GetErrorMessage(), _
            "종목코드 = " & iCode & ", row = " & row & ", col = " & col & ", reqtype = " & reqtype
        Exit Sub
    End If
    RQ_ONCE_VARIABLE(i, REQ_WRITE_ROW) = row
    RQ_ONCE_VARIABLE(i, REQ_WRITE_COL) = col
    RQ_ONCE_VARIABLE(i, REQ_TYPE) = reqtype
    
End Sub

Private Sub objShinhan_balance_ReceiveData(ByVal rqid As Integer)

    Dim i As Integer
    i = Sheets("v").Range("Selected_Account").Value2
    
    Dim nErr As Integer
    nErr = objShinhan_balance.GetErrorState()
    
    If nErr > 0 Then
        Receive_ErrorHandler "잔고조회", nErr, objShinhan_balance.GetErrorMessage(), CStr(i)
        Exit Sub
    End If
    
    On Error Resume Next
    Sheets("v").Range("Balance_" & i).FormulaR1C1 = CLng(objShinhan_balance.GetSingleData(12))
    
    '제비용단가 삭제
    With Sheets("Main")
        .Range(.Cells(StockReadStart, Autodeal_ownstock_col).Address, .Cells(StockReadFinish, Goodprice_deal_AUTO_Column).Address).ClearContents
        .Range(.Cells(FavoriteStart, Goodprice_deal_Favorite_Column).Address, .Cells(FavoriteFinish, Goodprice_deal_Favorite_Column).Address).ClearContents
    End With
    
    loadarray (LOAD_ARRAY_BALANCE)
    

    '주식보유시 제비용단가 입력
    Dim cnt As Integer, compareCode As String, goodprice As Long, owned_amount As Long
    Dim row As Integer, col As Integer, stockcode As String
    
    For cnt = 0 To CLng(objShinhan_balance.GetMultiRowCount) - 1
        
        goodprice = CLng(objShinhan_balance.GetMultiData(cnt, 12)) '12 : 제비용단가, 28 : 단축코드, 10 : 주문가능수량
        compareCode = objShinhan_balance.GetMultiData(cnt, 28) '12 : 제비용단가, 28 : 단축코드, 10 : 주문가능수량
        
        'Load stock first
        col = isReadCol
        For row = 1 To UBound(stockinfo, 1)
            stockcode = CStr(stockinfo(row, col))
            If Len(stockcode) = 6 And StrComp(stockcode, compareCode, vbBinaryCompare) = 0 Then 'row, column
                autodealinfo(row, iBalGoodPriceCol) = goodprice
                autodealinfo(row, iBalOwnAmountCol) = CLng(objShinhan_balance.GetMultiData(cnt, 10)) '주문가능수량
            End If
        Next row
        
        'Load Favorite
        col = iFReadCol
        For row = 1 To UBound(favoriteinfo, 1)
            stockcode = CStr(favoriteinfo(row, col))
            If Len(stockcode) = 6 And StrComp(stockcode, compareCode, vbBinaryCompare) = 0 Then   'row, column
                favoriteinfo(row, iFGoodPriceCol) = goodprice
            End If
        Next row
    Next

    'objv.Range("Balance_" & Account_number).FormulaR1C1 = Account_Balance
    Writearray (LOAD_ARRAY_FAVORITE)
    Writearray (LOAD_ARRAY_BALANCE)

End Sub

Private Sub objShinhan_buysell_ReceiveData(ByVal rqid As Integer)

    Dim i As Integer
    Dim nErr As Long

    '위치 구하기
    For i = 0 To RQSIZE
        If RQ_BUYSELL_VARIABLE(i, REQ_ID) = rqid Then
            RQ_BUYSELL_VARIABLE(i, REQ_ID) = REQ_INVALID_ID
            Exit For
        End If
    Next i
    If i = RQSIZE Then
        MsgBox "[구매/판매] 수신에 문제가 있습니다." & Chr(10) & "잠시후 다시 시도해주세요.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If

    nErr = objShinhan_buysell.GetErrorState()
    If nErr > 0 Then
        Receive_ErrorHandler "구매/판매", CInt(nErr), objShinhan_buysell.GetErrorMessage(), CStr(i)
        Exit Sub
    End If
    
    nErr = objShinhan_buysell.GetSingleData(0)
    If nErr = 0 Then
        Receive_ErrorHandler "구매/판매", 0, nErr & ": 계좌주문오류 !!", _
            i & "번, 부족금액/수량 : " & objShinhan_buysell.GetSingleData(3) & ", 가능수량 : " & objShinhan_buysell.GetSingleData(4) & ", 가능금액 : " & objShinhan_buysell.GetSingleData(5)
        Exit Sub
    End If
    
    Dim result As String
    Select Case RQ_BUYSELL_VARIABLE(i, REQ_AUTOSTOP)
    
        Case NOT_AUTOSTOP: If RQ_BUYSELL_VARIABLE(i, REQ_TYPE) = ORDER_BUY Then result = "매수중" Else result = "매도중"
        Case AUTOSTOP_GOOD: result = "익절중"
        Case AUTOSTOP_BAD: result = "손절중"
        
    End Select
    
    On Error Resume Next
    Sheets("Main").Cells(RQ_BUYSELL_VARIABLE(i, REQ_WRITE_ROW), RQ_BUYSELL_VARIABLE(i, REQ_WRITE_COL)).FormulaR1C1 = result
    
    'For i =
        
    ' RQ_BUYSELL_VARIABLE(i, REQ_TYPE) ' ORDER_BUY or ORDER_SELL
    ' RQ_BUYSELL_VARIABLE(i, REQ_STOCKCODE)
    ' RQ_BUYSELL_VARIABLE(i, REQ_PRICE) ' PRICE_MARKET_PRICE or price
    '
    '

End Sub

Private Sub objShinhan_ReceiveData(ByVal rqid As Integer)

    Dim i As Integer
    Dim nErr As Integer
    
    '위치 구하기
    For i = 0 To RQSIZE
        If RQ_ONCE_VARIABLE(i, REQ_ID) = rqid Then
            RQ_ONCE_VARIABLE(i, REQ_ID) = REQ_INVALID_ID
            Exit For
        End If
    Next i
    If i = RQSIZE Then
        MsgBox "[현재가] 수신에 문제가 있습니다." & Chr(10) & "잠시후 다시 시도해주세요.", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    nErr = objShinhan.GetErrorState()
    If nErr > 0 Then
        Receive_ErrorHandler "현재가", nErr, objShinhan.GetErrorMessage(), CStr(i)
        Exit Sub
    End If


    '실행
    On Error Resume Next
    Select Case RQ_ONCE_VARIABLE(i, REQ_TYPE)
        Case REQ_TYPE_PRICE
            Sheets("Main").Cells(RQ_ONCE_VARIABLE(i, REQ_WRITE_ROW), RQ_ONCE_VARIABLE(i, REQ_WRITE_COL)).FormulaR1C1 = CLng(objShinhan.GetSingleData(3))
        Case REQ_TYPE_NAME
            Sheets("Main").Cells(RQ_ONCE_VARIABLE(i, REQ_WRITE_ROW), RQ_ONCE_VARIABLE(i, REQ_WRITE_COL)).FormulaR1C1 = objShinhan.GetSingleData(5)
        Case REQ_TYPE_TRADEAMOUNT
            Sheets("Main").Cells(RQ_ONCE_VARIABLE(i, REQ_WRITE_ROW), RQ_ONCE_VARIABLE(i, REQ_WRITE_COL)).FormulaR1C1 = objShinhan.GetSingleData(7)
    End Select
    
    
End Sub

Private Sub RT_Price_Backend(ByVal rttype As Variant, ByRef i As Integer, ByRef nErr As Long, sErrMsg As String, ByRef price As String, ByRef tradeamount As String)

    If nErr > 0 Then
        Receive_ErrorHandler "실시간", 0, nErr & ": " & sErrMsg, CStr(i)
        Exit Sub
    End If
    
    Dim row As Integer
    With Sheets("Main")

        row = i + (RealtimeStart - StockReadStart + 1)
        stockinfo(row, isWriteCol) = price
        stockinfo(row, isAmountCol) = tradeamount
        '.Cells(RQ_CURPRICE_REALTIME(i, REQ_WRITE_ROW), RQ_CURPRICE_REALTIME(i, REQ_WRITE_COL)).FormulaR1C1 = price
        '.Cells(RQ_CURPRICE_REALTIME(i, REQ_WRITE_ROW), TradeAmountColumn).FormulaR1C1 = tradeamount
        
        '자동거래 마스터 스위치가 꺼져 있으면 종료
        If Sheets("v").Range("Autodeal_Enabled").Value2 = False Then Exit Sub
        '플러그인 사용시 자동거래 작동 종료
        If UsePlugin Then Exit Sub
        Dim goodprice As Long, thisprice As Long, owned_stock_amount As Long, idx As Integer
        idx = i + RealtimeStart
        '작동중이고, 거래가 끝나지 않았으며, 제비용단가와 보유량이 적힌 주식만 확인
        goodprice = .Cells(idx, Goodprice_deal_AUTO_Column).Value2
        owned_stock_amount = .Cells(idx, Autodeal_ownstock_col).Value2
        If .Cells(idx, Autodeal_CheckBox_Column).Value2 = True And .Cells(idx, Autodeal_status_col).Value2 = vbNullString And goodprice > 0 And owned_stock_amount > 0 Then
            
            '익절가가 있는데 익절가보다 높으면 거래
            thisprice = .Cells(idx, Autodeal_goodsell_Price_Col).Value2
            If thisprice > 0 And thisprice <= price Then
                SellStock CStr(.Cells(idx, StockReadColumn).Value2), PRICE_MARKET_PRICE, owned_stock_amount, idx, Autodeal_status_col, AUTOSTOP_GOOD
                .Cells(idx, Autodeal_CheckBox_Column).FormulaR1C1 = False
            End If
            '손절가가 있는데 손절가보다 낮으면 거래
            thisprice = .Cells(idx, Autodeal_badsell_Price_Col).Value2
            If thisprice > 0 And thisprice >= price Then
                SellStock CStr(.Cells(idx, StockReadColumn).Value2), PRICE_MARKET_PRICE, owned_stock_amount, idx, Autodeal_status_col, AUTOSTOP_BAD
                .Cells(idx, Autodeal_CheckBox_Column).FormulaR1C1 = False
            End If
            'RQ_CURPRICE_REALTIME(i, REQ_STOCKCODE)
            '.Cells(idx, StockReadColumn).Value2
        End If
        
    End With

End Sub

Private Sub oShin_favname_ReceiveData(ByVal rqid As Integer)
        
    req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) - 1
    Dim nErr As Integer
    nErr = oShin_favname.GetErrorState()
    If nErr > 0 Then
        Receive_ErrorHandler "현재가", oShin_favname.GetErrorState(), oShin_favname.GetErrorMessage(), "oShin_favname_ReceiveData"
        Exit Sub
    End If

    Dim rows As Integer
    For rows = 1 To UBound(favoriteinfo, 1)
        If favoriteinfo(rows, isNameCol) = CStr(rqid) Then
            favoriteinfo(rows, isNameCol) = oShin_favname.GetSingleData(5)
            Exit For
        End If
    Next rows
        
End Sub

Private Sub oShin_name_ReceiveData(ByVal rqid As Integer)
    
    req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) - 1
    Dim nErr As Integer
    nErr = oShin_name.GetErrorState()
    If nErr > 0 Then
        Receive_ErrorHandler "현재가", oShin_name.GetErrorState(), oShin_name.GetErrorMessage(), "oShin_name_ReceiveData"
        Exit Sub
    End If
    
    Dim rows As Integer
    For rows = 1 To UBound(stockinfo, 1)
        If stockinfo(rows, isNameCol) = CStr(rqid) Then
            stockinfo(rows, isNameCol) = oShin_name.GetSingleData(5)
            Exit For
        End If
    Next rows
    
End Sub

Private Sub oShin_PriceAmount_ReceiveData(ByVal rqid As Integer)

    req_remaining(LOAD_ARRAY_STOCKINFO) = req_remaining(LOAD_ARRAY_STOCKINFO) - 1
    Dim nErr As Integer
    nErr = oShin_PriceAmount.GetErrorState()
    If nErr > 0 Then
        Receive_ErrorHandler "현재가", oShin_PriceAmount.GetErrorState(), oShin_PriceAmount.GetErrorMessage(), "oShin_PriceAmount_ReceiveData"
        Exit Sub
    End If

    Dim rows As Integer
    For rows = 1 To UBound(stockinfo, 1)
        If stockinfo(rows, isWriteCol) = CStr(rqid) Then
            stockinfo(rows, isWriteCol) = oShin_PriceAmount.GetSingleData(3)
            stockinfo(rows, isAmountCol) = CStr(oShin_PriceAmount.GetSingleData(7))
            Exit For
        End If
    Next rows
    
End Sub

Private Sub oShin_PriceOnly_ReceiveData(ByVal rqid As Integer)

    req_remaining(LOAD_ARRAY_FAVORITE) = req_remaining(LOAD_ARRAY_FAVORITE) - 1
    
    Dim nErr As Integer
    nErr = oShin_PriceOnly.GetErrorState()
    If nErr > 0 Then
        Receive_ErrorHandler "현재가", oShin_PriceOnly.GetErrorState(), oShin_PriceOnly.GetErrorMessage(), "oShin_PriceOnly_ReceiveData"
        Exit Sub
    End If
    
    Dim rows As Integer
    For rows = 1 To UBound(favoriteinfo, 1)
        If favoriteinfo(rows, iFWriteCol) = CStr(rqid) Then
            favoriteinfo(rows, iFWriteCol) = oShin_PriceOnly.GetSingleData(3)
            Exit For
        End If
    Next rows
    
End Sub


Private Sub objShinhan_RT_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 0, objShinhan_RT.GetErrorState, objShinhan_RT.GetErrorMessage, CStr(objShinhan_RT.GetSingleData(3)), CStr(objShinhan_RT.GetSingleData(7))

End Sub

Private Sub objShinhan_RT2_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 1, objShinhan_RT2.GetErrorState, objShinhan_RT2.GetErrorMessage, objShinhan_RT2.GetSingleData(3), CStr(objShinhan_RT2.GetSingleData(7))

End Sub

Private Sub objShinhan_RT3_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 2, objShinhan_RT3.GetErrorState, objShinhan_RT3.GetErrorMessage, objShinhan_RT3.GetSingleData(3), CStr(objShinhan_RT3.GetSingleData(7))

End Sub

Private Sub objShinhan_RT4_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 3, objShinhan_RT4.GetErrorState, objShinhan_RT4.GetErrorMessage, objShinhan_RT4.GetSingleData(3), CStr(objShinhan_RT4.GetSingleData(7))

End Sub

Private Sub objShinhan_RT5_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 4, objShinhan_RT5.GetErrorState, objShinhan_RT5.GetErrorMessage, objShinhan_RT5.GetSingleData(3), CStr(objShinhan_RT5.GetSingleData(7))

End Sub

Private Sub objShinhan_RT6_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 5, objShinhan_RT6.GetErrorState, objShinhan_RT6.GetErrorMessage, objShinhan_RT6.GetSingleData(3), CStr(objShinhan_RT6.GetSingleData(7))

End Sub

Private Sub objShinhan_RT7_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 6, objShinhan_RT7.GetErrorState, objShinhan_RT7.GetErrorMessage, objShinhan_RT7.GetSingleData(3), CStr(objShinhan_RT7.GetSingleData(7))

End Sub

Private Sub objShinhan_RT8_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 7, objShinhan_RT8.GetErrorState, objShinhan_RT8.GetErrorMessage, objShinhan_RT8.GetSingleData(3), CStr(objShinhan_RT8.GetSingleData(7))

End Sub

Private Sub objShinhan_RT9_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 8, objShinhan_RT9.GetErrorState, objShinhan_RT9.GetErrorMessage, objShinhan_RT9.GetSingleData(3), CStr(objShinhan_RT9.GetSingleData(7))

End Sub

Private Sub objShinhan_RT10_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 9, objShinhan_RT10.GetErrorState, objShinhan_RT10.GetErrorMessage, objShinhan_RT10.GetSingleData(3), CStr(objShinhan_RT10.GetSingleData(7))

End Sub

Private Sub objShinhan_RT11_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 10, objShinhan_RT11.GetErrorState, objShinhan_RT11.GetErrorMessage, objShinhan_RT11.GetSingleData(3), CStr(objShinhan_RT11.GetSingleData(7))

End Sub

Private Sub objShinhan_RT12_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 11, objShinhan_RT12.GetErrorState, objShinhan_RT12.GetErrorMessage, objShinhan_RT12.GetSingleData(3), CStr(objShinhan_RT12.GetSingleData(7))

End Sub

Private Sub objShinhan_RT13_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 12, objShinhan_RT13.GetErrorState, objShinhan_RT13.GetErrorMessage, objShinhan_RT13.GetSingleData(3), CStr(objShinhan_RT13.GetSingleData(7))

End Sub

Private Sub objShinhan_RT14_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 13, objShinhan_RT14.GetErrorState, objShinhan_RT14.GetErrorMessage, objShinhan_RT14.GetSingleData(3), CStr(objShinhan_RT14.GetSingleData(7))

End Sub

Private Sub objShinhan_RT15_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 14, objShinhan_RT15.GetErrorState, objShinhan_RT15.GetErrorMessage, objShinhan_RT15.GetSingleData(3), CStr(objShinhan_RT15.GetSingleData(7))

End Sub

Private Sub objShinhan_RT16_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 15, objShinhan_RT16.GetErrorState, objShinhan_RT16.GetErrorMessage, objShinhan_RT16.GetSingleData(3), CStr(objShinhan_RT16.GetSingleData(7))

End Sub

Private Sub objShinhan_RT17_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 16, objShinhan_RT17.GetErrorState, objShinhan_RT17.GetErrorMessage, objShinhan_RT17.GetSingleData(3), CStr(objShinhan_RT17.GetSingleData(7))

End Sub

Private Sub objShinhan_RT18_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 17, objShinhan_RT18.GetErrorState, objShinhan_RT18.GetErrorMessage, objShinhan_RT18.GetSingleData(3), CStr(objShinhan_RT18.GetSingleData(7))

End Sub

Private Sub objShinhan_RT19_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 18, objShinhan_RT19.GetErrorState, objShinhan_RT19.GetErrorMessage, objShinhan_RT19.GetSingleData(3), CStr(objShinhan_RT19.GetSingleData(7))

End Sub

Private Sub objShinhan_RT20_ReceiveRTData(ByVal rttype As Variant)

    On Error Resume Next
    RT_Price_Backend rttype, 19, objShinhan_RT20.GetErrorState, objShinhan_RT20.GetErrorMessage, objShinhan_RT20.GetSingleData(3), CStr(objShinhan_RT20.GetSingleData(7))

End Sub

Public Sub MakeControl()
    'Dim ID As String    ' 멀티 접속용 로그인 ID
    Set oShin_PriceAmount = New GIExpertControl64
    Set oShin_PriceOnly = New GIExpertControl64
    Set oShin_name = New GIExpertControl64
    Set oShin_favname = New GIExpertControl64
    
    
    Set objShinhan = New GIExpertControl64          '조회 OCX
    Set objShinhan_buysell = New GIExpertControl64
    Set objShinhan_balance = New GIExpertControl64
    Set objShinhan_RT = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT2 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT3 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT4 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT5 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT6 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT7 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT8 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT9 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT10 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT11 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT12 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT13 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT14 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT15 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT16 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT17 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT18 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT19 = New GIExpertControl64          '조회 OCX
    Set objShinhan_RT20 = New GIExpertControl64          '조회 OCX
    'MsgBox objShinhan.StartIndi("redacted", "redacted", "redacted", "C:\SHINHAN-i\indi\giexpertstarter.exe")
    'ID = Cells(1, 6).Value
    'If Len(ID) > 0 Then
        'objShinhan.SetID (ID)
    'End If
End Sub

