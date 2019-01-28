VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose : Class stores values for ticker while looping through rows within worksheet


'- Class Variables
' Name of the ticker that the class contains information for
Private m_Ticker As String

' Total volume for the ticker; each day is added to this
Private m_TotalVolume As Double

' Open price of the ticker on the first day of year
Private m_FirstDayPrice As Double

' Close price of the ticker on the last day of the year
Private m_LastDayPrice As Double

' Difference between the FirstDayPrice and the LastDayPrice; calculated with the CompleteTicker
Private m_YearlyChangeActual As Double

' Percentage difference between the FirstDayPrice and the LastDayPrice; calculated with the CompleteTicker
Private m_YearlyChangePercentage As Double



'- Constructor
Private Sub Class_Initialize()
' Purpose : Creates a new instance and sets the default values

    m_Ticker = ""
    m_TotalVolume = 0
    m_FirstDayPrice = 0
    m_LastDayPrice = 0
End Sub



'- Properties
Public Property Get Ticker() As String
    Ticker = m_Ticker
End Property

Public Property Let Ticker(value As String)
    m_Ticker = value
End Property

Public Property Get FirstDayPrice() As Double
    FirstDayPrice = m_FirstDayPrice
End Property

Public Property Let FirstDayPrice(value As Double)
    m_FirstDayPrice = value
End Property

Public Property Get LastDayPrice() As Double
    LastDayPrice = m_LastDayPrice
End Property

Public Property Get TotalVolume() As Double
    TotalVolume = m_TotalVolume
End Property

Public Property Get YearlyChangeActual() As Double
    YearlyChangeActual = m_YearlyChangeActual
End Property

Public Property Get YearlyChangePercentage() As Double
    YearlyChangePercentage = m_YearlyChangePercentage
End Property



Public Sub AddVolume(dailyVolume As Double)
' Purpose : Updates the running total of the volume
'
' Accepts : dailyVolume     daily volume that is to be added to the existing volume

    m_TotalVolume = (m_TotalVolume + dailyVolume)

End Sub


Public Sub CompleteTicker(lastDayPriceValue As Double)
' Purpose : Calculates the
'
' Accepts : lastDayPriceValue   ticker price on the last day of the year

    m_LastDayPrice = lastDayPriceValue
    
    
    '- Calculate Yearly Change
    m_YearlyChangeActual = (m_LastDayPrice - m_FirstDayPrice)
    
    
    '- Calcualte Yearly Percentage
    '   Verify that the first day was not zero; when zero set yearly change to zero
    If (m_FirstDayPrice <> 0) Then
        m_YearlyChangePercentage = (m_YearlyChangeActual / m_FirstDayPrice)
    End If
    
End Sub

