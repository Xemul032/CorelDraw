Sub CalculatePerimeterAndArea()
    Dim s As Shape
    Dim newShape As Shape
    Dim perimeter As Double
    Dim area As Double
    Dim ratio As Double
    Dim originalUnits As cdrUnit
    Dim pageArea As Double
    Dim percentage As Double
    Dim koefKomp As Double
    Const maxArea = 4000
    Const plashkaArea = 400
    Const kk = 0.5
    Const procentZalivki = 20

    ' Сохранение текущих единиц измерения
    originalUnits = ActiveDocument.Unit

    ' Изменение единиц измерения на миллиметры
    ActiveDocument.Unit = cdrMillimeter

    ' Убедитесь, что выделено хотя бы одно изображение
    If ActiveSelection.Shapes.Count < 1 Then
        MsgBox "Пожалуйста, выделите хотя бы одну кривую или объект."
        ' Восстановление исходных единиц измерения
        ActiveDocument.Unit = originalUnits
        Exit Sub
    End If

    ' Если выделено более одного объекта, выводим сообщение
    If ActiveSelection.Shapes.Count > 1 Then
        MsgBox "Выделено несколько кривых! Выделите одну интересующую кривую!"
        ActiveDocument.Unit = originalUnits
        Exit Sub
    End If

    ' Выбор выделенного объекта
    Set s = ActiveSelection.Shapes(1)

    ' Проверяем, является ли объект кривой, если не является - преобразуем
    If s.Type <> cdrCurveShape Then
        s.ConvertToCurves
    End If

    ' Попытка разъединения объекта на кривые
    On Error Resume Next
    s.BreakApartEx
    On Error GoTo 0

    ' Проверка на количество оставшихся кривых
    If ActiveSelection.Shapes.Count > 1 Then
        MsgBox "После разъединения выделено несколько кривых! Пожалуйста, оставьте только одну."
        ActiveDocument.Unit = originalUnits
        Exit Sub
    ElseIf ActiveSelection.Shapes.Count = 0 Then
        MsgBox "После разъединения не осталось кривых! Проверьте объект."
        ActiveDocument.Unit = originalUnits
        Exit Sub
    End If
    ' Вычисляем периметр
    perimeter = s.Curve.Length

    ' Вычисляем площадь
    area = s.Curve.area

    ' Вычисляем коэффициент компактности
    koefKomp = (4 * 3.1415 * area) / (perimeter * perimeter)
    koefKomp1 = perimeter / (4 * Sqr(area))

    ' Вычисляем площадь страницы
    pageArea = ActivePage.SizeWidth * ActivePage.SizeHeight

    ' Вычисляем процент от площади страницы
    percentage = (area / pageArea) * 100

    

    ' Проверяем площадь с максимумом и минимумом
    If area < maxArea And area > plashkaArea Then
        If kk > koefKomp Then
            MsgBox "Нанесение проходит!" & vbCrLf & _
                   "Процент MGI: " & Format(percentage, "0.00") & "%" & vbCrLf & _
                   "KoefKompakt: " & Format(koefKomp, "0.00000") & vbCrLf & _
                   "KoefKompakt1: " & Format(koefKomp1, "0.00000") & vbCrLf & _
                   "Площадь MGI: " & Format(area, "0.00") & " квадратных мм"
        ElseIf procentZalivki > percentage Then
            MsgBox "Нанесение проходит!" & vbCrLf & _
                    "KoefKompakt: " & Format(koefKomp, "0.00000") & vbCrLf & _
                    "KoefKompakt1: " & Format(koefKomp1, "0.00000") & vbCrLf & _
                   "Процент MGI: " & Format(percentage, "0.00") & "%" & vbCrLf & _
                   "Площадь MGI: " & Format(area, "0.00") & " квадратных мм"
        Else
            MsgBox "Нанесение НЕ проходит!" & vbCrLf & _
                   "Процент MGI: " & Format(percentage, "0.00") & "%" & vbCrLf & _
                   "Площадь MGI: " & Format(area, "0.00") & " квадратных мм"
        End If
    End If
    
' Проверяем площадь с максимумом
    If area > maxArea Then
    
     If kk > koefKomp Then
            MsgBox "Нанесение проходит!" & vbCrLf & _
                   "Процент MGI: " & Format(percentage, "0.00") & "%" & vbCrLf & _
                   "KoefKompakt1: " & Format(koefKomp1, "0.00000") & vbCrLf & _
                   "KoefKompakt: " & Format(koefKomp, "0.00000") & vbCrLf & _
                   "Площадь MGI: " & Format(area, "0.00") & " квадратных мм" & perimeter
                   
                   Else
    
        MsgBox "Площадь MGI нанесения превышает допустимый порог! Не берем в работу!" & vbCrLf & _
                "KoefKompakt: " & Format(koefKomp, "0.00000") & vbCrLf & _
                "KoefKompakt1: " & Format(koefKomp1, "0.00000") & vbCrLf & _
               "Процент MGI: " & Format(percentage, "0.00") & "%" & vbCrLf & _
               "Площадь MGI: " & Format(area, "0.00") & " квадратных мм"
        ActiveDocument.Unit = originalUnits
        Exit Sub
        End If
    End If
    
    ' Проверяем площадь с минимумом
    If area < plashkaArea Then
        MsgBox "Нанесение проходит!" & vbCrLf & _
               "Процент MGI: " & Format(percentage, "0.00") & "%" & vbCrLf & _
               "Площадь MGI: " & Format(area, "0.00") & " квадратных мм"
    End If

    ' Вычисляем соотношение периметра к площади
    If perimeter <> 0 Then
        ratio = area / perimeter
    Else
        MsgBox "Периметр равен нулю, возможно, кривая не замкнута."
        ActiveDocument.Unit = originalUnits
        Exit Sub
    End If

    ' Форматирование результатов с двумя знаками после запятой
    Dim formattedPerimeter As String
    Dim formattedArea As String
    Dim formattedRatio As String

    formattedPerimeter = Format(perimeter, "0.00")
    formattedArea = Format(area, "0.00")
    formattedRatio = Format(ratio, "0.00")

    ' Восстановление исходных единиц измерения
    ActiveDocument.Unit = originalUnits

    
End Sub
