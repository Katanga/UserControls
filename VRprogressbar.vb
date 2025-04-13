Imports System.Windows.Forms
Imports System.Drawing

Public Class VRprogressbar
  Inherits ProgressBar

	Private _label As String = ""
  Private _percent As Integer = 0
  Public Property Label As String
    Get
      Return _label
    End Get
    Set(value As String)
      _label = value
      Me.Invalidate() ' forza il ridisegno
    End Set
  End Property

  Public Sub New()
    ' Imposta lo stile per consentire il ridisegno
    Me.SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
  End Sub

  ''' <summary>
  ''' Gestisco la proprietà Value per aggiungere il calcolo della percentuale
  ''' </summary>
  Public Shadows Property Value As Integer
    Get
      Return MyBase.Value
    End Get
    Set(value As Integer)
      If MyBase.Value <> value Then
        MyBase.Value = value
        CalcolaPercentuale()
        CalcolaRemainingTime()
      End If
    End Set
  End Property
  ''' <summary>
  ''' Gestisco la proprietà Minimum per aggiungere il calcolo della percentuale
  ''' </summary>
  Public Shadows Property Minimum As Integer
    Get
      Return MyBase.Minimum
    End Get
    Set(value As Integer)
      If MyBase.Minimum <> value Then
        MyBase.Minimum = value
        CalcolaPercentuale()
      End If
    End Set
  End Property
  ''' <summary>
  ''' Gestisco la proprietà Maximum per aggiungere il calcolo della percentuale
  ''' </summary>
  Public Shadows Property Maximum As Integer
    Get
      Return MyBase.Maximum
    End Get
    Set(value As Integer)
      If MyBase.Maximum <> value Then
        MyBase.Maximum = value
        CalcolaPercentuale()
      End If
    End Set
  End Property
	''' <summary>
	''' Restituisco la percentuale calcolata, 
	''' se imposto la percentuale ricalcolo value
	''' </summary>
	Public Property Percent As Integer
		Get
			Return _percent
		End Get
		Set(value As Integer)
			_percent = value
			MyBase.Value = CInt((value / 100) * Me.Maximum)
		End Set
	End Property
  ''' <summary>
  ''' Colore usato per riempire la barra
  ''' </summary>
  Public Shadows Property BackColor As Color
    Get
      Return MyBase.BackColor
    End Get
    Set(value As Color)
      MyBase.BackColor = value
      BrushBar = New SolidBrush(value)
    End Set
  End Property

  Public Shadows Property ForeColor As Color
    Get
      Return MyBase.ForeColor
    End Get
    Set(value As Color)
      MyBase.ForeColor = value
      BrushText = New SolidBrush(value)
    End Set
  End Property

  ''' <summary>
  ''' Brush usato per disegnare la barra
  ''' </summary>
  Private BrushBar As SolidBrush = Brushes.LightGreen
  ''' <summary>
  ''' Colore usato per disegnare il testo
  ''' </summary>
  Private BrushText As SolidBrush = Brushes.Black

  ''' <summary>
  ''' Istante Zero da cui calcolare il tempo trascorso e la stima a finire
  ''' </summary>
  Public Property StartTime As DateTime = DateTime.Now
  Public Property Estimated As TimeSpan
  Public Property Remaining As TimeSpan
  ''' <summary>
  ''' Mostra la % sopra la progressBar
  ''' </summary>
  Public Property ShowPercent As Boolean = True
  ''' <summary>
  ''' Mostra il tempo rimanente sopra la progressbar
  ''' </summary>
  Public Property ShowRemaining As Boolean = True
  ''' <summary>
  ''' testo che precede il tempo rimanente
  ''' </summary>
  Public Property LabelRemaining As String = ""

  Private Sub CalcolaRemainingTime()
    Dim elapsed As TimeSpan = DateTime.Now - StartTime
    Dim ratio As Double = Me.Value / Me.Maximum
    If ratio > 0 Then
      Estimated = TimeSpan.FromTicks(CLng(elapsed.Ticks / ratio))
      Remaining = Estimated - elapsed
      Dim minuti As Integer = CInt(Remaining.TotalMinutes)
      Dim secondi As Integer = Remaining.Seconds
      ' Il tempo restante, se previsto, è visulizzato a partire dal 5% per evitare stime troppo sbagliate
      _label = IIf(ShowPercent, $"{Percent:0}% ", "") & IIf(ShowRemaining And Percent >= 5, $" {LabelRemaining} {minuti}:{secondi:00}", "")
    End If
  End Sub

  Private Sub CalcolaPercentuale()
    If Me.Maximum > 0 Then
      _percent = CInt((Me.Value / Me.Maximum) * 100)
    Else
      _percent = 0
    End If
  End Sub

  Protected Overrides Sub OnPaint(e As PaintEventArgs)
    MyBase.OnPaint(e)

    If MyBase.Visible Then
      ' Disegna lo sfondo base
      Dim rect As Rectangle = Me.ClientRectangle
      If ProgressBarRenderer.IsSupported Then
        ProgressBarRenderer.DrawHorizontalBar(e.Graphics, rect)
      End If

      ' Calcola la larghezza della parte "piena"
      Dim filledRect As New Rectangle(0, 0, CInt(rect.Width * (Me.Value / Me.Maximum)), rect.Height)
      e.Graphics.FillRectangle(BrushBar, filledRect)
      ' Brushes.LightGreen

      ' Disegna il testo centrato
      If Not String.IsNullOrEmpty(Label) Then
        Dim sf As New StringFormat With {
                .Alignment = StringAlignment.Center,
                .LineAlignment = StringAlignment.Center
            }
        e.Graphics.DrawString(Label, MyBase.Font, BrushText, rect, sf)
      End If
    End If
  End Sub
End Class

