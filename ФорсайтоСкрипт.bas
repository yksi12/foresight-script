Attribute VB_Name = "ФорсайтоСкрипт"
' ФорсайтоСкрипт
' Версия 0.1
' Copyright (c) 2013 Юрий Куприянов
' Данная лицензия разрешает лицам, получившим копию данного программного обеспечения и сопутствующей документации
' (в дальнейшем именуемыми «Программное Обеспечение»), безвозмездно использовать Программное Обеспечение без ограничений,
' включая неограниченное право на использование, копирование, изменение, слияние, публикацию, распространение,
' сублицензирование и/или продажу копий Программного Обеспечения, а также лицам, которым предоставляется данное
' Программное Обеспечение, при соблюдении следующих условий:
' Указанное выше уведомление об авторском праве и данные условия должны быть включены во все копии или
' значимые части данного Программного Обеспечения.
' ДАННОЕ ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ ПРЕДОСТАВЛЯЕТСЯ «КАК ЕСТЬ», БЕЗ КАКИХ-ЛИБО ГАРАНТИЙ, ЯВНО ВЫРАЖЕННЫХ ИЛИ ПОДРАЗУМЕВАЕМЫХ,
' ВКЛЮЧАЯ ГАРАНТИИ ТОВАРНОЙ ПРИГОДНОСТИ, СООТВЕТСТВИЯ ПО ЕГО КОНКРЕТНОМУ НАЗНАЧЕНИЮ И ОТСУТСТВИЯ НАРУШЕНИЙ, НО НЕ ОГРАНИЧИВАЯСЬ ИМИ.
' НИ В КАКОМ СЛУЧАЕ АВТОРЫ ИЛИ ПРАВООБЛАДАТЕЛИ НЕ НЕСУТ ОТВЕТСТВЕННОСТИ ПО КАКИМ-ЛИБО ИСКАМ, ЗА УЩЕРБ ИЛИ ПО ИНЫМ ТРЕБОВАНИЯМ,
' В ТОМ ЧИСЛЕ, ПРИ ДЕЙСТВИИ КОНТРАКТА, ДЕЛИКТЕ ИЛИ ИНОЙ СИТУАЦИИ, ВОЗНИКШИМ ИЗ-ЗА ИСПОЛЬЗОВАНИЯ ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ
' ИЛИ ИНЫХ ДЕЙСТВИЙ С ПРОГРАММНЫМ ОБЕСПЕЧЕНИЕМ.


Public Enum CardTypes
    ctTrend = 2 'trends are red
    ctSubTrend = 1 'subtrends are pink?
    ctTechnology = 4 'technologies are blue
    ctFormat = 3 'formats are green
    ctEvent = 4 'events are orange?
    ctPolicy = 6 'policies are violet
    ctPossibility = 7 'possibilities are orange
    ctThreat = 0 'threats are grey
    ctMarket = 5 'markets are yellow
    ctUnknown = 255
End Enum

Sub GenerateForesightMap()
'
' Главный макрос макрос
'

' Запускаем Visio
Dim vis As Visio.Application
Set vis = openVisio

' Делаем тренды
Dim Trends As New Collection
Dim i, k
k = 0
i = 2
With Application.ActiveSheet
While .Cells(i, 1).Value <> ""
    Dim shpColor
    Dim shpMove
    shpMove = 0
    
    Dim shpType As CardTypes

    If .Cells(i, 2).Value = "тренд" Or .Cells(i, 2).Value = "подтренд" Then
        shpType = ctTrend
        shpColorText = "RGB(255,0,0)" 'trends are red
        shpColor = RGB(255, 0, 0) 'trends are red
        shpMove = -3
    ElseIf .Cells(i, 2).Value = "подтренд" Then
        shpType = ctSubTrend
        shpColorText = "RGB(255,50,50)" 'trends are red
        shpColor = RGB(255, 50, 50) 'trends are red
        shpMove = -2.8
    ElseIf .Cells(i, 2).Value = "формат" Then
        shpType = ctFormat
        shpColorText = "RGB(0,255,0)" 'formats are green
        shpColor = RGB(0, 255, 0) 'formats are green
    ElseIf .Cells(i, 2).Value = "технология" Then
        shpType = ctTechnology
        shpColorText = "RGB(0,0,255)" 'technologies are blue
        shpColor = RGB(0, 0, 255) 'technologies are blue
    ElseIf .Cells(i, 2).Value = "возможность" Then
        shpType = ctPossibility
        shpColorText = "RGB(255,165,0)" 'possibilities are orange
        shpColor = RGB(255, 165, 0) 'possibilities are orange
    ElseIf .Cells(i, 2).Value = "угроза" Then
        shpType = ctThreat
        shpColorText = "RGB(100,100,100)" 'threats are grey
        shpColor = RGB(100, 100, 100) 'threats are grey
    ElseIf .Cells(i, 2).Value = "нормативный акт" Then
        shpType = ctPolicy
        shpColorText = "RGB(185, 40, 170)" 'policies are violet
        shpColor = RGB(185, 40, 170) 'policies are violet
    Else
        shpType = ctUnknown
        shpColorText = "RGB(160,160,160)" 'others are grey
        shpColor = RGB(160, 160, 160) 'others are grey
    End If

        Dim shp As Visio.Shape
        Set shp = vis.ActiveWindow.Page.Shapes.ItemFromID(36).Duplicate
    
        Dim shpLink As String
        Dim parentTrend As String
        If .Cells(i, 9).Text <> "" Then
            shpLink = .Cells(i, 9).Text
            parentTrend = Split(.Cells(i, 9).Text, ",")(0)
        End If

        ' Вытаскиваем поля карточки
        Dim cardNo, cardType, cardTitle, cardText, cardYear As String
        cardNo = .Cells(i, 1).Text
        cardType = UCase(.Cells(i, 2).Text)
        cardTitle = UCase(Left(.Cells(i, 3).Text, 1)) + Right(.Cells(i, 3).Text, Len(.Cells(i, 3).Text) - 1)
        
        If .Cells(i, 5).Text <> "" Then
            cardText = UCase(Left(.Cells(i, 5).Text, 1)) + Right(.Cells(i, 5).Text, Len(.Cells(i, 5).Text) - 1)
        Else
            cardText = .Cells(i, 5).Text
        End If
        cardYear = .Cells(i, 6).Text

        'Добавляем текст карточки
        If shpType = ctTrend Then
            ShapeText = cardTitle + Chr(10) + cardText + Chr(10) + cardYear + "| #" + cardNo
        ElseIf shpType = ctSubTrend Then
            ShapeText = cardType + Chr(10) + cardTitle + Chr(10) + cardText + Chr(10) + cardYear + "| #" + cardNo + "|базовый тренд " + shpLink
        Else
            ShapeText = cardType + Chr(10) + cardTitle + Chr(10) + cardText + Chr(10) + cardYear + "| #" + cardNo + "|связано с " + shpLink
        End If

        Dim vsoCharacters1 As Visio.Characters
        Set vsoCharacters1 = shp.Characters
        vsoCharacters1.Begin = 0
        vsoCharacters1.End = 0
        vsoCharacters1.Text = ShapeText

        shp.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = shpColorText

        Set momPosX = vis.ActiveWindow.Page.Shapes.ItemFromID(36).Cells("PinX")
        Set momPosY = vis.ActiveWindow.Page.Shapes.ItemFromID(36).Cells("PinY")

        Set celObj = shp.Cells("height")
        localCent = celObj.Result("inches")
        ShapeHeight = Format(localCent, "000.0000")

        If Not shpType = ctTrend And Not shpType = ctSubTrend Then
            vsoCharacters1.Begin = 0
            vsoCharacters1.End = Len(cardType)
            vsoCharacters1.CharProps(visCharacterStyle) = 17#
            vsoCharacters1.CharProps(visCharacterColor) = shpType

            If IsNumeric(.Cells(i, 6).Text) Then
                yr = (.Cells(i, 6).Value - 2015) * 1.2
                shpMove = yr
            End If

            PosX = yr + 3
            PosY = momPosY
            On Error Resume Next
            PosY = Trends.Item(Trim(Left(parentTrend, 2)))
            On Error GoTo 0

            'dy = k * (ShapeHeight + 0.2)
        Else
            PosX = 1.7
            PosY = momPosY - k * (ShapeHeight + 0.2)
            k = k + 1
            Trends.Add PosY, CStr(cardNo)
        End If

        shp.Cells("PinX") = PosX
        shp.Cells("PinY") = PosY


'        With vis
'            .ActiveWindow.DeselectAll
'            .ActiveWindow.Select shp, visSelect
'
'            .Application.ActiveWindow.Selection.Move shpMove, -dy
'        End With
    
    i = i + 1
Wend

        With vis
            .ActiveWindow.DeselectAll
            .ActiveWindow.Select .ActiveWindow.Page.Shapes.ItemFromID(36), visSelect

            .Application.ActiveWindow.Selection.Delete
    End With

End With

End Sub

Function openVisio() As Object
Dim vis As Visio.Application
Set vis = CreateObject("Visio.Application")
With vis
    .Visible = True
    .Documents.Open Application.ActiveWorkbook.Path + "\foresight-map.vsd"
End With

Set openVisio = vis

End Function

