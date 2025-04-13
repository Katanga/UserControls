Imports ADODB

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class RichTextBoxVR
	Inherits System.Windows.Forms.RichTextBox

	'UserControl esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Richiesto da Progettazione Windows Form
	Private components As System.ComponentModel.IContainer

	'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
	'Può essere modificata in Progettazione Windows Form.  
	'Non modificarla mediante l'editor del codice.
	<System.Diagnostics.DebuggerStepThrough()>
	Private Sub InitializeComponent()
		components = New System.ComponentModel.Container()
	End Sub
	''' <summary>
	''' Metodo personalizzato per aggiungere del testo con un colore 
	''' e poi ripristinare il colore originale
	''' </summary>
	''' <param name="text">Testo da scrivere</param>
	''' <param name="InkColor">Colore inchiostro da utilizzare, se non specificato usa il colore corrente</param>
	''' <param name="Refresh">True forza il refresh della RichtextBox</param>
	Public Sub AppendTextColored(text As String, Optional ByVal InkColor As Color = Nothing, Optional Refresh As Boolean = False)
		Me.SuspendLayout()
		Me.SelectionStart = Me.Text.Length
		Dim oldcolor As Color = Me.SelectionColor
		If IsNothing(InkColor) Then
			' Clore non specificato, uso quello corrente
			InkColor = oldcolor
		End If
		If oldcolor <> InkColor Then
			' il colore è diverso dall'attuale, lo uso e poi ripristino il vecchio
			Me.SelectionColor = InkColor
			Me.AppendText(text)
			Me.SelectionColor = oldcolor
		Else
			Me.AppendText(text)
		End If
		Me.ScrollToCaret()
		Me.ResumeLayout()
		If Refresh Then
			Me.Refresh()
		End If
	End Sub
	''' <summary>
	''' Metodo personalizzato per aggiungere del testo + fine linea con un colore 
	''' </summary>
	''' <param name="text">Testo da scrivere</param>
	''' <param name="InkColor">Colore inchiostro da utilizzare</param>
	''' <param name="Refresh">True forza il refresh della RichtextBox</param>
	Public Sub AppendLine(text As String, Optional InkColor As Color = Nothing, Optional Refresh As Boolean = True)
		AppendTextColored(text & vbCrLf, InkColor, Refresh)
	End Sub

End Class
