Option Explicit On
Option Strict On
Imports AxGIEXPERTCONTROLLib
Imports System.Threading

Public Class CRT_And_Autodeal

    Private WithEvents RealtimeHandler As AxGiExpertControl
    Private WithEvents BuySellHandler As AxGiExpertControl

    Public Structure Posinfo
        Public rows As Integer
        Public col_name As Integer
        Public col_price As Integer
        Public col_amount As Integer
    End Structure
    Private pos As Posinfo
    Private rtcode As String
    Private status As Boolean
    Private ret_name As Short
    Private ret_priceamount As Short

    Private Const ORDER_BUY As Integer = 2
    Private Const ORDER_SELL As Integer = 1

    Private Const PRICE_MARKET_PRICE As Integer = -1
    Private Const STOCK_ALL As Integer = -255

    Private Const NONE_RELATED As Integer = 0
    Private Const BADSELL As Integer = 1
    Private Const GOODSELL As Integer = 2

    Public Sub New()
        RealtimeHandler = New AxGiExpertControl
        RealtimeHandler.CreateControl()
        BuySellHandler = New AxGiExpertControl
        BuySellHandler.CreateControl()
        status = False
    End Sub

    Public Sub RT_Start(ByRef sCode As String, ByRef adr As Posinfo) '종목코드, 수정변수, row, col
        pos = adr
        rtcode = sCode

        RealtimeHandler.SetQueryName("SC")
        RealtimeHandler.SetSingleData(0, CStr(rtcode))
        ret_priceamount = RealtimeHandler.RequestData()
        If ret_priceamount <= 0 Then
            MsgBox("[realtime] 송신오류 : " & RealtimeHandler.GetErrorState() & Chr(10) & CStr(RealtimeHandler.GetErrorMessage()) & Chr(10) & "SC, Code=" & ret_priceamount, CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        RealtimeHandler.SetQueryName("SB")
        RealtimeHandler.SetSingleData(0, CStr(rtcode))
        ret_name = RealtimeHandler.RequestData()
        If ret_name <= 0 Then
            MsgBox("[realtime] 송신오류 : " & RealtimeHandler.GetErrorState() & Chr(10) & CStr(RealtimeHandler.GetErrorMessage()) & Chr(10) & "SB, Code=" & ret_name, CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
            Exit Sub
        End If

        status = RealtimeHandler.RequestRTReg("SC", CStr(rtcode))
        If status = False Then MsgBox("[realtime] 송신오류 : " & RealtimeHandler.GetErrorState() & Chr(10) & CStr(RealtimeHandler.GetErrorMessage()) & Chr(10) & "rtreg, Code=" & rtcode, CType(vbOKOnly + vbExclamation, MsgBoxStyle), "PowerVBA")
        'Dim thread As New Thread(AddressOf Form1.Thread_for_updategrid)
        'Thread.Start(pos.rows)
    End Sub
    Public Sub RT_Stop()
        RealtimeHandler.UnRequestRTReg("SC", CStr(rtcode))
        status = False
        'Dim thread As New Thread(AddressOf Form1.Thread_for_updategrid)
        'Thread.Start(pos.rows)
    End Sub

    Public Function Enabled() As Boolean
        Enabled = status
    End Function

    Public Sub AssignDoctrine() '조건 입력 (계좌, 보유주수, 제비용단가, 손절가, 퍼센트, 익절가, 퍼센트)

    End Sub

    Public Sub Dispose() 'RECEIVERTDATA 중단, 데이터 제거
        RT_Stop()
        BuySellHandler.SelfMemFree(True)
        BuySellHandler.Dispose()
        RealtimeHandler.SelfMemFree(True)
        RealtimeHandler.Dispose()
    End Sub

    Private Sub BuyStock(ByRef price As Long, ByRef amount As Integer)
        Buysell_backend(ORDER_BUY, price, amount, NONE_RELATED)
    End Sub

    Private Sub SellStock(ByRef price As Long, ByRef amount As Integer, ByRef isgood As Integer)
        Buysell_backend(ORDER_SELL, price, amount, isgood)
    End Sub

    Private Sub Buysell_backend(ByRef buysell As Integer, ByRef price As Long, ByRef amount As Integer, ByRef isgood As Integer)

        With Form1

            Call BuySellHandler.SetQueryName("SABA101U1")
            Call BuySellHandler.SetSingleData(0, .AccountNumString)
            Call BuySellHandler.SetSingleData(1, "redacted")
            Call BuySellHandler.SetSingleData(2, "redacted")
            Call BuySellHandler.SetSingleData(3, vbNullString)
            Call BuySellHandler.SetSingleData(4, vbNullString)
            Call BuySellHandler.SetSingleData(5, "0")
            Call BuySellHandler.SetSingleData(6, "00")
            Call BuySellHandler.SetSingleData(7, CStr(buysell))
            Call BuySellHandler.SetSingleData(8, "A" & rtcode)
            If amount = STOCK_ALL Then Call BuySellHandler.SetSingleData(9, CStr(.autodealinfo(pos.rows, .iAutoOwnCol))) Else Call BuySellHandler.SetSingleData(9, amount)
            If price = PRICE_MARKET_PRICE Then Call BuySellHandler.SetSingleData(10, "0") Else Call BuySellHandler.SetSingleData(10, price)
            Call BuySellHandler.SetSingleData(11, "1")
            If price = PRICE_MARKET_PRICE Then Call BuySellHandler.SetSingleData(12, "1") Else Call BuySellHandler.SetSingleData(12, "2")
            Call BuySellHandler.SetSingleData(13, "0")
            Call BuySellHandler.SetSingleData(14, "0")
            Call BuySellHandler.SetSingleData(15, vbNullString)
            Call BuySellHandler.SetSingleData(16, vbNullString)
            Call BuySellHandler.SetSingleData(21, "Y")

            ''매수매도
            Dim shit As Short = BuySellHandler.RequestData
            If shit <= 0 Then
                .Handle_Error("[CRITICAL] 자동거래 송신오류 : " & BuySellHandler.GetErrorState() & ", " & CStr(BuySellHandler.GetErrorMessage()) & ", " _
                                & "buysell = " & buysell & "종목코드 = " & rtcode & ", 가격 = " & price & ", 주 = " & amount & ", r = " & pos.rows + .StockReadStart - 1)
                Exit Sub
            End If
            .autodealinfo(pos.rows, .iAutoStatusCol) = shit
            .autodealinfo(pos.rows, .iAutoBoxCol) = False
            Select Case isgood
                Case GOODSELL : .Handle_Error("[DEBUG] 익절요청 완료, 행 : " & pos.rows + .StockReadStart - 1)
                Case BADSELL : .Handle_Error("[DEBUG] 손절요청 완료, 행 : " & pos.rows + .StockReadStart - 1)
            End Select

        End With

    End Sub

    Private Sub BuySellHandler_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles BuySellHandler.ReceiveData

        With Form1
            Dim nErr As Short = BuySellHandler.GetErrorState()
            If nErr > 0 Then
                .Handle_Error("[CRITICAL] 익절/손절요청 수신중 오류발생 : " & BuySellHandler.GetErrorState() & ", " & CStr(BuySellHandler.GetErrorMessage()) & Chr(10) & "행 :" & pos.rows + .StockReadStart - 1)
                Exit Sub
            End If

            nErr = CShort(BuySellHandler.GetSingleData(0))
            If nErr = 0 Then
                .Handle_Error("[CRITICAL] 익절/손절 실패 : " & pos.rows + .StockReadStart - 1 & "번, 부족금액/수량 : " & CStr(BuySellHandler.GetSingleData(3)) & ", 가능수량 : " _
                & CStr(BuySellHandler.GetSingleData(4)) & ", 가능금액 : " & CStr(BuySellHandler.GetSingleData(5)))
                .autodealinfo(pos.rows, .iAutoStatusCol) = "오류"
                Exit Sub
            End If
            '주문번호를 입력
            .autodealinfo(pos.rows, .iAutoStatusCol) = CStr(BuySellHandler.GetSingleData(0))
            .Handle_Error("[DEBUG] 익절/손절요청이 승인되었습니다. 행 : " & pos.rows + .StockReadStart - 1)

            .Invoke(New Form1.WriteGrid(AddressOf Form1.UpdateGrid), New Object() {pos.rows})
            .Invoke(New Form1.Write_Onbehalf(AddressOf Form1.Writearray), New Object() {4}) ' 4: LOAD_ARRAY_AUTODEAL
        End With

    End Sub

    Private Sub RealtimeHandler_ReceiveData(sender As Object, e As _DGiExpertControlEvents_ReceiveDataEvent) Handles RealtimeHandler.ReceiveData

        With Form1
            Dim nErr As Integer = RealtimeHandler.GetErrorState()
            If nErr > 0 Then
                .Handle_Error("[WARNING] 실시간 수신오류 : " & RealtimeHandler.GetErrorState() & ", " & CStr(RealtimeHandler.GetErrorMessage()) & Chr(10) & "행 :" & pos.rows + .StockReadStart - 1)
                Exit Sub
            End If

            Select Case e.rqid
                Case ret_name
                    .stockinfo(pos.rows, pos.col_name) = RealtimeHandler.GetSingleData(5)
                Case ret_priceamount
                    .stockinfo(pos.rows, pos.col_price) = RealtimeHandler.GetSingleData(3)
                    .stockinfo(pos.rows, pos.col_amount) = RealtimeHandler.GetSingleData(7)
            End Select

            .Invoke(New Form1.WriteGrid(AddressOf Form1.UpdateGrid), New Object() {pos.rows})
        End With

    End Sub

    Private Sub RealtimeHandler_ReceiveRTData(sender As Object, e As _DGiExpertControlEvents_ReceiveRTDataEvent) Handles RealtimeHandler.ReceiveRTData

        With Form1

            Dim nErr As Integer = RealtimeHandler.GetErrorState()
            If nErr > 0 Then
                .Handle_Error("[WARNING] 실시간 수신오류 : " & RealtimeHandler.GetErrorState() & ", " & CStr(RealtimeHandler.GetErrorMessage()) & Chr(10) & "행 :" & pos.rows + .StockReadStart - 1)
                Exit Sub
            End If

            .stockinfo(pos.rows, pos.col_price) = RealtimeHandler.GetSingleData(3)
            .stockinfo(pos.rows, pos.col_amount) = RealtimeHandler.GetSingleData(7)
            .Invoke(New Form1.WriteGrid(AddressOf Form1.UpdateGrid), New Object() {pos.rows})

            '마스터키가 꺼져 있으면 작동금지
            If .AUTODEAL_ENABLED_MASTERKEY = False Then Exit Sub
            '로컬키가 꺼져 있으면 작동금지
            If CBool(.autodealinfo(pos.rows, .iAutoBoxCol)) = False Then Exit Sub
            '이미 끝난 거래인 경우 작동금지
            If CStr(.autodealinfo(pos.rows, .iAutoStatusCol)) <> vbNullString Then Exit Sub

            ' 제비용단가와 보유량이 제대로 적힌 주식만 확인
            Dim goodprice As Long, owned_stock_amount As Long
            Try
                goodprice = CLng(.autodealinfo(pos.rows, .iBalGoodPriceCol))
                owned_stock_amount = CLng(.autodealinfo(pos.rows, .iAutoOwnCol))
            Catch
                .Handle_Error("[CRITICAL] 제비용단가나 보유주식에 잘못된 값이 입력되었습니다. 행 : " & pos.rows + .StockReadStart - 1)
                Exit Sub
            End Try
            If goodprice <= 0 Or owned_stock_amount <= 0 Then
                .Handle_Error("[CRITICAL] 제비용단가나 보유주식이 입력되지 않았습니다. 행 : " & pos.rows + .StockReadStart - 1)
                Exit Sub
            End If

            '익절가, 손절가가 제대로 쓰여있는지 확인후 거래 (익절 -> 손절 순)
            Dim currentprice As Long
            Dim thisprice As Long
            currentprice = CLng(RealtimeHandler.GetSingleData(3))

            '익절가가 쓰여있는 경우 익절
            If CStr(.autodealinfo(pos.rows, .iAutoGoodsellCol)) <> vbNullString Then
                Try
                    thisprice = CLng(.autodealinfo(pos.rows, .iAutoGoodsellCol))
                Catch ex As Exception
                    .Handle_Error("[CRITICAL] 익절가가 올바르지 않습니다. 행 : " & pos.rows + .StockReadStart - 1)
                    Exit Sub
                End Try
                If thisprice < 0 Then
                    .Handle_Error("[CRITICAL] 익절가가 음수입니다. 행 : " & pos.rows + .StockReadStart - 1)
                    Exit Sub
                End If

                If thisprice <= currentprice Then
                    SellStock(PRICE_MARKET_PRICE, STOCK_ALL, GOODSELL)
                    'writearray needed
                End If
            End If

            '손절가가 쓰여있는 경우 손절
            If CStr(.autodealinfo(pos.rows, .iAutoBadsellCol)) <> vbNullString Then
                Try
                    thisprice = CLng(.autodealinfo(pos.rows, .iAutoBadsellCol))
                Catch ex As Exception
                    .Handle_Error("[CRITICAL] 손절가가 올바르지 않습니다. 행 : " & pos.rows + .StockReadStart - 1)
                    Exit Sub
                End Try
                If thisprice < 0 Then
                    .Handle_Error("[CRITICAL] 손절가가 음수입니다. 행 : " & pos.rows + .StockReadStart - 1)
                    Exit Sub
                End If

                If thisprice >= currentprice Then
                    SellStock(PRICE_MARKET_PRICE, STOCK_ALL, BADSELL)
                    'writearray needed
                    '.Invoke(New .Write_Onbehalf(AddressOf Form1.Writearray), New Object() { .LOAD_ARRAY_AUTODEAL})
                End If
            End If

        End With

    End Sub

End Class
