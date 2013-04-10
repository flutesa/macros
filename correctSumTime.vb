Sub BurkovaAS()
'
' BurkovaAS Макрос
'

    Dim i, sum_min, sum_hour, hour_from_min As Integer
    Dim time, min, hour As String
    
    sum_min = 0
    sum_hour = 0
    
    ReDim time(2) As String
    
    For i = 3 To Cells(Rows.Count, "J").End(xlUp).Row
        time = Split(Cells(i, "J").Value, ":")
        hour = time(0)
        min = time(1)
        
        sum_hour = sum_hour + CStr(hour)
        sum_min = sum_min + CStr(min)
    Next i
    
    'сумма по времени "как есть"
    hour = Trim(Str(sum_hour))
    min = Trim(Str(sum_min))
    
    Cells(1, "J").Value = hour + ":" + min
    
    'сумма по времени преобразованная
    hour_from_min = sum_min \ 60 'сколько из получившейся суммы минут можем получить дополнительно часов часов
    sum_hour = sum_hour + hour_from_min 'к имеющейся сумме часов прибавляем дополнительные часы, полученные из минут
    sum_min = sum_min Mod 60 'сколько минут у нас осталось в итоге
    
    hour = Trim(Str(sum_hour))
    min = Trim(Str(sum_min))
    
    Cells(2, "K").Value = hour + ":" + min
    
End Sub
