### Макросы преобразования значения INT в ip адрес

```
Function DEC_TO_IP(ByVal DecNumber As String) As String
    ' Функция преобразует строку с большим числом в IP-адрес
    Dim Result As String
    Dim Octet As Long
    Dim i As Integer
    Dim Num As Double ' Используем Double как промежуточное значение
    Dim Temp As String
    
    ' Проверяем, является ли строка числом
    If Not IsNumeric(DecNumber) Then
        DEC_TO_IP = "Ошибка: не число"
        Exit Function
    End If
    
    ' Преобразуем строку в число (Double имеет достаточную точность для 32-бит)
    Num = CDbl(DecNumber)
    
    ' Если число слишком велико для 32-битного IP
    If Num > 4294967295# Or Num < 0 Then
        DEC_TO_IP = "Ошибка: не 32-бит"
        Exit Function
    End If
    
    ' Основной алгоритм преобразования
    Result = ""
    For i = 1 To 4
        ' Выделяем октет: делим на 256^(4-i) и берём целую часть
        Octet = Int(Num / (256 ^ (4 - i)))
        ' Вычитаем выделенный октет из исходного числа
        Num = Num - Octet * (256 ^ (4 - i))
        
        ' Формируем строку с октетами
        If Result <> "" Then
            Result = Result & "." & CStr(Octet)
        Else
            Result = CStr(Octet)
        End If
    Next i
    
    DEC_TO_IP = Result
End Function
```

### Макросы преобразования ip адреса в INT значение
```
Function IP_TO_DEC(ByVal IP_Address As String) As String
    ' Функция преобразует IP-адрес в строку с десятичным числом
    Dim i As Integer
    Dim Result As Double
    Dim Parts
    
    ' Проверяем корректность формата
    If IP_Address = "" Then
        IP_TO_DEC = "#ОШИБКА: пустая строка"
        Exit Function
    End If
    
    ' Разбиваем строку на части по точке
    Parts = Split(IP_Address, ".")
    
    ' Проверяем количество октетов
    If UBound(Parts) <> 3 Then
        IP_TO_DEC = "#ОШИБКА: неверный формат IP"
        Exit Function
    End If
    
    ' Инициализируем результат
    Result = 0
    
    ' Основной расчёт с проверкой каждого октета
    For i = 0 To 3
        ' Проверяем, является ли октет числом
        If Not IsNumeric(Parts(i)) Then
            IP_TO_DEC = "#ОШИБКА: октет " & (i+1) & " не число"
            Exit Function
        End If
        
        Dim octetValue As Integer
        octetValue = CInt(Parts(i))
        
        ' Проверяем диапазон (0-255)
        If octetValue < 0 Or octetValue > 255 Then
            IP_TO_DEC = "#ОШИБКА: октет " & (i+1) & " вне диапазона"
            Exit Function
        End If
        
        ' Вычисляем: o1*256³ + o2*256² + o3*256 + o4
        Select Case i
            Case 0: Result = Result + octetValue * 16777216  ' 256^3
            Case 1: Result = Result + octetValue * 65536     ' 256^2
            Case 2: Result = Result + octetValue * 256       ' 256^1
            Case 3: Result = Result + octetValue             ' 256^0
        End Select
    Next i
    
    ' Возвращаем результат как строку
    IP_TO_DEC = CStr(Result)
End Function
```
