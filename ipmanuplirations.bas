'�A�h���X�d���`�F�b�N
Function CheckAddressDuplicate(ByVal str1 As String, ByVal str2 As String) As Boolean
  If str1 = str2 Then
    CheckAddressDuplicate = True
  Else
      CheckAddressDuplicate = False
  End If
End Function

'IP�A�h���X�̌`���`�F�b�N
'�A�h���X���������؂���ꍇ�Fmode��"address"
'�}�X�N���������؂���ꍇ�@�Fmode��"mask"
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

'�A�h���X�������W�������`�F�b�N����
Function IsAddressInRange(ByVal addr As String, ByVal str As String) As Boolean
  Dim network As String
  Dim numNetwork As Long
  Dim mask As Integer, numMask As Double, numMaskInv As Double, numAddr As Double
  Dim head As Double, tail As Double
  
  ' �l�b�g�}�X�N��32bit�\�����v�Z
  mask = CInt(ExtractAddress(str, "mask"))
  numMask = Mask2Bits(mask)
  numMaskInv = invertBits(numMask, 32)
  
    
  numNetwork = IPaddr2Num(ExtractAddress(str, "host"))
  head = numNetwork And numMask    '�擪�A�h���X(�l�b�g���[�N)���Z�o
  tail = numNetwork Or numMaskInv  '�����A�h���X(�u���[�h�L���X�g)���Z�o
  
  ' �ΏۃA�h���X�̐��l�\�����v�Z
  numAddr = IPaddr2Num(addr)
  
  IsAddressInRange = False
  
  If head > tail Then
    If (head >= numAddr) And (tail <= numAddr) Then IsAddressInRange = True
  Else
    If (head <= numAddr) And (tail >= numAddr) Then IsAddressInRange = True
  End If
      
End Function

'�A�h���X�����ꃌ���W�������`�F�b�N����
Function chkHostsInSameNetwork(ByVal hostA As String, hostB As String, network As String) As Boolean
  chkHostsInSameNetwork = False
  chkHostsInSameNetwork = IsAddressInRange(hostA, network) And IsAddressInRange(hostB, network)
End Function

'�A�h���X/�}�X�N�\�L�̃A�h���X��񂩂�A�h���X/�}�X�N�̂����ꂩ�𒊏o����
Function ExtractAddress(ByVal str As String, mode As String) As String
  Dim Host As String, mask As String
  ExtractAddress = ""
  If mode = "mask" Then
    ExtractAddress = Mid(str, InStr(str, "/") + 1)
  ElseIf mode = "host" Then
    ExtractAddress = Left(str, InStr(str, "/") - 1)
  End If
End Function

'IP�A�h���X��Long�^�����ɕϊ�����
Function IPaddr2Num(ByVal str As String) As Long
  '������
  IPaddr2Num = 0
  Dim octets As Variant
  Dim tmp As String
  If CheckAddressFormat(str, "address") Then
    '�I�N�e�b�g���Ƃɕ���
    octets = Split(str, ".")
    For i = 0 To UBound(octets)
      octets(i) = Right(("0" & Hex(CInt(octets(i)))), 2)
    Next i
  
    '���o�����e�I�N�e�b�g��16�i���ɕϊ����Đڑ�
    tmp = "&H"
    tmp = tmp & octets(0) & octets(1) & octets(2) & octets(3)
  
    '16�i���\����IP�A�h���X������𐔒l�^��
    IPaddr2Num = CLng(tmp)
    
  End If
End Function

'Long�^������IP�A�h���X�ɕϊ�����
Function Num2IPaddr(ByVal adr As Long) As String
  '������
  Dim tmp As String
  Num2IPaddr = ""
  
  '16�i�\���ɕϊ�����
  tmp = CStr(Hex(adr))
  
  '�擪��0������8�P�^�Ƀg���~���O����
  '��2�̕␔�Ƃ��Đ擪�ɂ�"FF"���������邽��10�P�^�ɂ��Ă���8�P�^��
  While (Len(tmp) < 10)
    tmp = "0" & tmp
  Wend
  tmp = Right(tmp, 8)
  
  '�e�I�N�e�b�g������
  Num2IPaddr = CStr(CInt("&H" & Mid(tmp, 1, 2)))
  Num2IPaddr = Num2IPaddr & "." & CStr(CInt("&H" & Mid(tmp, 3, 2)))
  Num2IPaddr = Num2IPaddr & "." & CStr(CInt("&H" & Mid(tmp, 5, 2)))
  Num2IPaddr = Num2IPaddr & "." & CStr(CInt("&H" & Mid(tmp, 7, 2)))

End Function

'�w��̒����̃r�b�g�}�X�N�𐶐�����
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

'�w��̒����̃��C���h�J�[�h�}�X�N�𐶐�����
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

'�S�r�b�g�𔽓]�����Ďw��̃r�b�g���ɂ��낦��
Function invertBits(ByVal x As Long, bits As Integer) As Long
  If (x And &H40000000) = 0 Then
    invertBits = x Xor ((2 ^ bits) - 1)
  Else
    x = x And &H7FFFFFFF
    invertBits = x Xor &H7FFFFFFF
  End If
End Function

'�n�~���O����(���Ⴕ�Ă���r�b�g��)���v�Z����
Function getHummingDistance(ByVal a As Long, b As Long) As Integer
  Dim x As Long     '�r�b�g��r���ʊi�[�p
  
  x = a Xor b       '�قȂ��Ă���bit�����o
  
  getHummingDistance = 0
  Dim i As Integer
  '�����Ă���r�b�g�̐��𐔂���(31�r�b�g��)
  For i = 0 To 30
    If (x And (2 ^ i)) > 0 Then getHummingDistance = getHummingDistance + 1
  Next i
  '�擪�r�b�g�������Ă��邱�Ƃ𔻒肷��
  If (x And &H80000000) Then getHummingDistance = getHummingDistance + 1

End Function


'�@�\�@�F����NW���̔C�ӂ�IP�A�h���X�Ԃɂ����郏�C���h�J�[�h����������
'�@�\�A�F����NW��2��IP�A�h���X�ԂɔC�ӂ̒����̃��C���h�J�[�h���\���ł��邩���肵�A
'�@�@�@�@�\���\�ȃ��C���h�J�[�h��ԋp����
Function generateWildcard(ByVal head As Long, tail As Long, network As String, _
    Optional x As Integer, Optional maxAddr As Long) As String()
  
  '�ϐ��������@
  Dim tmp() As String   '�����ς݃��C���h�J�[�h�}�X�N�i�[�p
  Dim flg As Boolean    '���샂�[�h����(�@�\�@���A��)
  flg = False
  If x = 0 Then x = 1   '�@�\�@�̂Ƃ��̓}�X�N���͎w�肵�Ȃ��̂Ńf�t�H���g�l���Z�b�g
  If maxAddr = 0 Then   '�@�\�@�̂Ƃ���
    maxAddr = tail      '   �ő�A�h���X�͎w�肵�Ȃ��̂Ńf�t�H���g�l�Ƃ��ďI���A�h���X���Z�b�g
    flg = True          '   ���샂�[�h����t���O���Z�b�g
  End If
  
  ReDim tmp(1) As String
  tmp(0) = "n/a"
  generateWildcard = tmp
  
  '�J�n�A�h���X�ƏI���A�h���X�������NW�ɑ��݂���Ȃ珈�����J�n
  If chkHostsInSameNetwork(Num2IPaddr(head), Num2IPaddr(tail), network) Then
    'Debug.Print "examing " & Num2IPaddr(head) & "->" & Num2IPaddr(tail)
    'Debug.Print "with flg=" & CStr(flg) & ", maxAddr=" & Num2IPaddr(maxAddr)

    '�ϐ��������A
    Dim i As Long           '�����J�n�A�h���X
    Dim nextAddr(3) As Long '�����ΏۃA�h���X�i�[�p
    Dim n As Integer        '�����ς݃��C���h�J�[�h���̕ۑ��p(�z��J�E���^)
    Dim m As Integer        '���C���h�J�[�h�}�X�N���ۑ��p
    Dim k As Integer        '(2^�}�X�N��)�͈̔͂Ɋ܂܂��A�h���X�̌����i�[
    n = 0
    i = head
    m = x
    k = 2 ^ m
    
    '�@�\�A�Ƃ��ČĂ΂ꂽ�ꍇ�A�����Ώە���ύX����
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
      
      '������m[bits]�̃��C���h�J�[�h���\���ł��邩���肷��
      '�����P�F�ΏۃA�h���X�ƌ����͈͏I���A�h���X�̊Ԃ�m�r�b�g���̃A�h���X������
      '�����Q�F�ΏۃA�h���X�ƌ����͈͏I���A�h���X�̃n�~���O������m[bits]�ł���
      'Debug.Print "h(" & Num2IPaddr(i) & "->" & Num2IPaddr(i + x) & ")=" & getHummingDistance(i, i + x)
      If (i <= (tail - x)) And (getHummingDistance(i, i + x) = m) Then
        
        nextAddr(0) = i
        nextAddr(1) = i + x
        nextAddr(2) = nextAddr(1) + 1
        nextAddr(3) = nextAddr(2) + x
        
        '������m+1[bits]�̃��C���h�J�[�h�}�X�N���\���ł��邩���s����
        '�����P�F�g�������A�h���X�͈͂��w�肳��Ă���A�h���X���𒴂��Ă��Ȃ�
        '�����Q�F��������(m�r�b�g)�̃��C���h�J�[�h�ш�Ɨאڂ��Ă���
        '�����R�F2�̃��C���h�J�[�h�ш�ɂ����ėאڂ��Ă���2�A�h���X��
        '�@�@�@�@�n�~���O������m+1�r�b�g�Ȃ烏�C���h�J�[�h�𓝍�����
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
          '���C���h�J�[�h�̊g�����ł��Ȃ�������m�r�b�g�̂��̂Ŋm��
          tmp(n) = Num2IPaddr(i) & " " & Num2IPaddr(mask2Wildcard(m))
          i = i + k
        End If
        'Debug.Print "recirsive_level=" & nest & " n=" & n & " " & tmp(n)
      Else '���C���h�J�[�h���\���ł��Ȃ��ꍇ��host(0.0.0.0)�w��Ŋm��
        tmp(n) = "host " & Num2IPaddr(i)
        'Debug.Print "n=" & n & " " & tmp(n)
        i = i + 1
      End If
      n = n + 1
    Wend
    generateWildcard = tmp
  End If
  
End Function





