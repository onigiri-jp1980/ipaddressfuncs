'アドレス重複チェック
Function CheckAddressDuplicate(ByVal str1 As String, ByVal str2 As String) As Boolean
  If str1 = str2 Then
    CheckAddressDuplicate = True
  Else
      CheckAddressDuplicate = False
  End If
End Function

'IPアドレスの形式チェック
'アドレスだけを検証する場合：mode→"address"
'マスクだけを検証する場合　：mode→"mask"
Function CheckAddressFormat(ByVal str1 As String, Optional ByVal mode As String) As Boolean
  Dim Reg
  Dim Ptn As String
  Set Reg = CreateObject("VBScript.RegExp")
  Ptn = "(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])/([0-9]|[1-2][0-9]|3[0-2])$"
  If mode = "address" Then
    Ptn = "(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$"
  ElseIf mode = "mask" Then
    Ptn = "([0-9]|[1-2][0-9]|3[0-2])$"
  End If
  With Reg
    .Global = True
    .IgnoreCase = True
    .Pattern = Ptn
    If (.test(str1)) Then
      CheckAddressFormat = True
      Else
        CheckAddressFormat = False
    End If
  End With
End Function

'アドレスがレンジ内かをチェックする
Function IsAddressInRange(ByVal addr As String, ByVal str As String) As Boolean
  Dim network As String
  Dim numNetwork As Long
  Dim mask As Integer, numMask As Double, numMaskInv As Double, numAddr As Double
  Dim head As Double, tail As Double
  
  ' ネットマスクの32bit表現を計算
  mask = CInt(ExtractAddress(str, "mask"))
  numMask = Mask2Bits(mask)
  numMaskInv = invertBits(numMask, 32)
  
    
  numNetwork = IPaddr2Num(ExtractAddress(str, "host"))
  head = numNetwork And numMask    '先頭アドレス(ネットワーク)を算出
  tail = numNetwork Or numMaskInv  '末尾アドレス(ブロードキャスト)を算出
  
  ' 対象アドレスの数値表現を計算
  numAddr = IPaddr2Num(addr)
  
  IsAddressInRange = False
  
  If head > tail Then
    If (head >= numAddr) And (tail <= numAddr) Then IsAddressInRange = True
  Else
    If (head <= numAddr) And (tail >= numAddr) Then IsAddressInRange = True
  End If
      
End Function

'アドレスが同一レンジ内かをチェックする
Function chkHostsInSameNetwork(ByVal hostA As String, hostB As String, network As String) As Boolean
  chkHostsInSameNetwork = False
  chkHostsInSameNetwork = IsAddressInRange(hostA, network) And IsAddressInRange(hostB, network)
End Function

'アドレス/マスク表記のアドレス情報からアドレス/マスクのいずれかを抽出する
Function ExtractAddress(ByVal str As String, mode As String) As String
  Dim Host As String, mask As String
  ExtractAddress = ""
  If mode = "mask" Then
    ExtractAddress = Mid(str, InStr(str, "/") + 1)
  ElseIf mode = "host" Then
    ExtractAddress = Left(str, InStr(str, "/") - 1)
  End If
End Function

'IPアドレスをLong型整数に変換する
Function IPaddr2Num(ByVal str As String) As Long
  '初期化
  IPaddr2Num = 0
  Dim octets As Variant
  Dim tmp As String
  If CheckAddressFormat(str, "address") Then
    'オクテットごとに分解
    octets = Split(str, ".")
    For i = 0 To UBound(octets)
      octets(i) = Right(("0" & Hex(CInt(octets(i)))), 2)
    Next i
  
    '取り出した各オクテットを16進数に変換して接続
    tmp = "&H"
    tmp = tmp & octets(0) & octets(1) & octets(2) & octets(3)
  
    '16進数表現のIPアドレス文字列を数値型に
    IPaddr2Num = CLng(tmp)
    
  End If
End Function

'Long型整数をIPアドレスに変換する
Function Num2IPaddr(ByVal adr As Long) As String
  '初期化
  Dim tmp As String
  Num2IPaddr = ""
  
  '16進表現に変換する
  tmp = CStr(Hex(adr))
  
  '先頭に0を補って8ケタにトリミングする
  '※2の補数として先頭につく"FF"を処理するため10ケタにしてから8ケタに
  While (Len(tmp) < 10)
    tmp = "0" & tmp
  Wend
  tmp = Right(tmp, 8)
  
  '各オクテットを処理
  Num2IPaddr = CStr(CInt("&H" & Mid(tmp, 1, 2)))
  Num2IPaddr = Num2IPaddr & "." & CStr(CInt("&H" & Mid(tmp, 3, 2)))
  Num2IPaddr = Num2IPaddr & "." & CStr(CInt("&H" & Mid(tmp, 5, 2)))
  Num2IPaddr = Num2IPaddr & "." & CStr(CInt("&H" & Mid(tmp, 7, 2)))

End Function

'指定の長さのビットマスクを生成する
Function Mask2Bits(ByVal mask As Integer) As Long
  If mask < 31 Then
    Dim tmp As Long
    tmp = &HC0000000
      For i = 1 To mask - 2
      tmp = tmp Or (2 ^ (30 - i))
    Next i
    Mask2Bits = tmp
  ElseIf mask = 31 Then
    Mask2Bits = &HFFFFFFFC
  ElseIf mask = 32 Then
    Mask2Bits = &HFFFFFFFF
  End If
  
End Function

'指定の長さのワイルドカードマスクを生成する
Function mask2Wildcard(ByVal mask As Integer) As Long
  Dim i
  mask2Wildcard = 0
  If mask = 32 Then
    mask2Wildcard = &HFFFFFFFF
  ElseIf mask = 31 Then
    mask2Wildcard = &H7FFFFFFF
  Else
    For i = 0 To mask - 1
      mask2Wildcard = mask2Wildcard Or (2 ^ i)
    Next i
  End If
End Function

'全ビットを反転させて指定のビット長にそろえる
Function invertBits(ByVal x As Long, bits As Integer) As Long
  If (x And &H40000000) = 0 Then
    invertBits = x Xor ((2 ^ bits) - 1)
  Else
    x = x And &H7FFFFFFF
    invertBits = x Xor &H7FFFFFFF
  End If
End Function

'ハミング距離(相違しているビット数)を計算する
Function getHummingDistance(ByVal a As Long, b As Long) As Integer
  Dim x As Long     'ビット比較結果格納用
  
  x = a Xor b       '異なっているbitを検出
  
  getHummingDistance = 0
  Dim i As Integer
  '立っているビットの数を数える(31ビット分)
  For i = 0 To 30
    If (x And (2 ^ i)) > 0 Then getHummingDistance = getHummingDistance + 1
  Next i
  '先頭ビットが立っていることを判定する
  If (x And &H80000000) Then getHummingDistance = getHummingDistance + 1

End Function


'機能①：同一NW内の任意のIPアドレス間におけるワイルドカードを検索する
'機能②：同一NWの2つのIPアドレス間に任意の長さのワイルドカードが構成できるか判定し、
'　　　　構成可能なワイルドカードを返却する
Function generateWildcard(ByVal head As Long, tail As Long, network As String, _
    Optional x As Integer, Optional maxAddr As Long) As String()
  
  '変数初期化①
  Dim tmp() As String   '生成済みワイルドカードマスク格納用
  Dim flg As Boolean    '動作モード判定(機能①か②か)
  flg = False
  If x = 0 Then x = 1   '機能①のときはマスク長は指定しないのでデフォルト値をセット
  If maxAddr = 0 Then   '機能①のときは
    maxAddr = tail      '   最大アドレスは指定しないのでデフォルト値として終了アドレスをセット
    flg = True          '   動作モード判定フラグをセット
  End If
  
  ReDim tmp(1) As String
  tmp(0) = "n/a"
  generateWildcard = tmp
  
  '開始アドレスと終了アドレスが同一のNWに存在するなら処理を開始
  If chkHostsInSameNetwork(Num2IPaddr(head), Num2IPaddr(tail), network) Then
    'Debug.Print "examing " & Num2IPaddr(head) & "->" & Num2IPaddr(tail)
    'Debug.Print "with flg=" & CStr(flg) & ", maxAddr=" & Num2IPaddr(maxAddr)

    '変数初期化②
    Dim i As Long           '検索開始アドレス
    Dim nextAddr(3) As Long '検索対象アドレス格納用
    Dim n As Integer        '生成済みワイルドカード数の保存用(配列カウンタ)
    Dim m As Integer        'ワイルドカードマスク長保存用
    Dim k As Integer        '(2^マスク長)の範囲に含まれるアドレスの個数を格納
    n = 0
    i = head
    m = x
    k = 2 ^ m
    
    '機能②として呼ばれた場合、検索対象幅を変更する
    If flg = False Then
      x = (2 ^ x) - 1
      'nest = m
    Else
      x = 2 ^ (x - 1)
    End If
    'Debug.Print "m=" & CStr(m) & " x=" & CStr(x) & " k=" & CStr(k)
    
   
    While i <= tail
      'Debug.Print ("adr=" & Num2IPaddr(i))
      ReDim Preserve tmp(n) As String
      
      '長さがm[bits]のワイルドカードが構成できるか判定する
      '条件１：対象アドレスと検索範囲終了アドレスの間にmビット幅のアドレスがある
      '条件２：対象アドレスと検索範囲終了アドレスのハミング距離がm[bits]である
      'Debug.Print "h(" & Num2IPaddr(i) & "->" & Num2IPaddr(i + x) & ")=" & getHummingDistance(i, i + x)
      If (i <= (tail - x)) And (getHummingDistance(i, i + x) = m) Then
        
        nextAddr(0) = i
        nextAddr(1) = i + x
        nextAddr(2) = nextAddr(1) + 1
        nextAddr(3) = nextAddr(2) + x
        
        '長さがm+1[bits]のワイルドカードマスクが構成できるか試行する
        '条件１：拡張したアドレス範囲が指定されているアドレス幅を超えていない
        '条件２：同じ長さ(mビット)のワイルドカード帯域と隣接している
        '条件３：2つのワイルドカード帯域において隣接している2つアドレスの
        '　　　　ハミング距離がm+1ビットならワイルドカードを統合する
        If (nextAddr(3) <= maxAddr) And _
          (getHummingDistance(nextAddr(0), nextAddr(1)) = getHummingDistance(nextAddr(2), nextAddr(3))) And _
          (getHummingDistance(nextAddr(1), nextAddr(2)) = getHummingDistance(nextAddr(0), nextAddr(1)) + 1) _
        Then
          Dim tmp2() As String
          'Debug.Print "h("; Num2IPaddr(i) & "->" & Num2IPaddr(nextAddr(3)) & " is extendable"
          tmp2 = generateWildcard(i, nextAddr(3), network, m + 1, maxAddr)
          tmp(n) = tmp2(0)
          i = i + IPaddr2Num(Split(tmp(n), " ")(1)) + 1
        Else
          'ワイルドカードの拡張ができなかったらmビットのもので確定
          tmp(n) = Num2IPaddr(i) & " " & Num2IPaddr(mask2Wildcard(m))
          i = i + k
        End If
        'Debug.Print "recirsive_level=" & nest & " n=" & n & " " & tmp(n)
      Else 'ワイルドカードを構成できない場合はhost(0.0.0.0)指定で確定
        tmp(n) = "host " & Num2IPaddr(i)
        'Debug.Print "n=" & n & " " & tmp(n)
        i = i + 1
      End If
      n = n + 1
    Wend
    generateWildcard = tmp
  End If
  
End Function





