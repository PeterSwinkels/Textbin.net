'The imports and settings used by this module.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic.ControlChars
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Convert
Imports System.Environment
Imports System.Linq

'This module contains this program's core procedures.
Public Module TextBinDemoModule
   'The data types supported by this demo.
   Private Enum DataTypesE As Integer
      None = -1           'Nothing.
      DLLReferences       'Anything that could be a valid DLL filename.
      EMailAddresses      'Anything that could be a valid e-mail address.
      GuidIds             'Anything that could be a valid GUID id.
      HumanReadable       'Any string of "human readable" characters (character codes 31-127.)
      Names               'Anything that could be a name (last, initials (first).)
      URLs                'Anything that could be a valid url (with the protocol specified.)
   End Enum

   'The relative text fragment positions checked by this demo.
   Private Enum RelativePositionsE As Integer
      RPNone = 0     'No position.
      RPStart = 1    'The start position.
      RPMiddle = 2   'The middle position.
      RPEnd = 4      'The end position.
   End Enum

   Private WithEvents TextBin As New TextBinClass   'Contains a reference to the TextBin class.

   'This procedure checks whether specified characters occur the specified number of times.
   Private Function CheckCounts(Text As String, Counts As String) As Boolean
      Try
         Dim Character As Char = Nothing
         Dim Count As Integer = Nothing
         Dim SubItems As New List(Of String)

         For Each Item As String In SplitText(Counts, ","c, "\"c)
            SubItems = SplitText(Item, ":"c, "\"c)
            Character = SubItems.First().First()
            If SubItems(1).StartsWith("<") Then
               Count = ToInt32(SubItems(1).Substring(1))
               If Not Text.Split(Character).Length - 1 < Count Then Return False
            ElseIf SubItems(1).StartsWith(">") Then
               Count = ToInt32(SubItems(1).Substring(1))
               If Not Text.Split(Character).Length - 1 > Count Then Return False
            Else
               Count = ToInt32(SubItems(1))
               If Not Text.Split(Character).Length - 1 = Count Then Return False
            End If
         Next Item
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return True
   End Function

   'This procedure checks whether or not the specified fragment occurs at the specified positions.
   Private Function CheckFragmentPositions(Text As String, Fragment As String, Positions As RelativePositionsE, Optional CaseSensitive As Boolean = True, Optional ExpectedResult As Boolean = True) As Boolean
      Try
         Dim FoundPositions As RelativePositionsE = RelativePositionsE.RPNone

         If Not CaseSensitive Then
            Fragment = Fragment.ToLower()
            Text = Text.ToLower()
         End If

         If Text = Fragment Then
            FoundPositions = (RelativePositionsE.RPStart Or RelativePositionsE.RPMiddle Or RelativePositionsE.RPEnd)
         Else
            If Text.StartsWith(Fragment) Then FoundPositions = FoundPositions Or RelativePositionsE.RPStart
            If Text.EndsWith(Fragment) Then FoundPositions = FoundPositions Or RelativePositionsE.RPEnd
            If Text.IndexOf(Fragment, 1) > 0 AndAlso Text.IndexOf(Fragment) < Text.Length - Fragment.Length Then FoundPositions = FoundPositions Or RelativePositionsE.RPMiddle
         End If

         Return ((FoundPositions = Positions) = ExpectedResult)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure gives the command to check whether or not the specified fragments occur at the specified positions.
   Private Function CheckFragmentsPositions(Text As String, FragmentsPositions As String) As Boolean
      Try
         Dim Arguments As String = Nothing
         Dim CaseSensitive As Boolean = True
         Dim Character As Char = Nothing
         Dim ExpectedResult As Boolean = True
         Dim Fragment As String = Nothing
         Dim Positions As RelativePositionsE = RelativePositionsE.RPNone
         Dim SubItems As New List(Of String)

         For Each Item As String In SplitText(FragmentsPositions, ","c, "\"c)
            SubItems = SplitText(Item, ":"c, "\"c)
            Fragment = SubItems.First()
            Arguments = SubItems(1).ToUpper()
            CaseSensitive = Not Arguments.Contains("I")
            ExpectedResult = Not Arguments.Contains("F")
            Positions = RelativePositionsE.RPNone
            If Arguments.Contains("E") Then Positions = Positions Or RelativePositionsE.RPEnd
            If Arguments.Contains("M") Then Positions = Positions Or RelativePositionsE.RPMiddle
            If Arguments.Contains("S") Then Positions = Positions Or RelativePositionsE.RPStart
            If Not CheckFragmentPositions(Text, Fragment, Positions, CaseSensitive, ExpectedResult) Then Return False
         Next Item

         Return True
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the used to further filter the search results.
   Private Function CurrentDataType(Optional NewDataType As DataTypesE = DataTypesE.None, Optional NewIncludeUnicode As Boolean = Nothing) As DataTypesE
      Try
         Static DataType As DataTypesE = DataTypesE.HumanReadable

         If Not NewDataType = DataTypesE.None Then
            DataType = NewDataType

            Select Case DataType
               Case DataTypesE.DLLReferences
                  TextBin.DefineText(ToInt32(" "c), ToInt32("~"c), , "\/:*?""<>|", NewIncludeUnicode)
               Case DataTypesE.EMailAddresses
                  TextBin.DefineText(ToInt32("!"c), ToInt32("~"c), , "()[]\;:,<>""", NewIncludeUnicode)
               Case DataTypesE.GuidIds
                  TextBin.DefineText(ToInt32("0"c), ToInt32("9"c), "ABCDEFabcdef-{}", , NewIncludeUnicode)
               Case DataTypesE.HumanReadable
                  TextBin.DefineText(ToInt32(" "c), ToInt32("~"c), Cr & Tab, , NewIncludeUnicode)
               Case DataTypesE.Names
                  TextBin.DefineText(ToInt32("A"c), ToInt32("Z"c), "abcdefghijklmnopqrstuvwxyz(,.) ", , NewIncludeUnicode)
               Case DataTypesE.URLs
                  TextBin.DefineText(ToInt32("!"c), ToInt32("~"c), , "<>""'", NewIncludeUnicode)
               Case Else
                  TextBin.DefineText(ToInt32(" "c), ToInt32("~"c), Cr & Tab, , NewIncludeUnicode)
                  Console.WriteLine("Invalid data type! Using default data type.")
            End Select
         End If

         Return DataType
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return DataTypesE.None
   End Function

   'This procedure display's this program's information.
   Private Sub DisplayInformation()
      Try
         With My.Application.Info
            Console.WriteLine($"{ .Title} v{ .Version} - by: { .CompanyName}")
            Console.WriteLine()
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure indicates whether the specified text fits the current data type.
   Private Function FitsDataType(Text As String) As Boolean
      Try
         Select Case CurrentDataType()
            Case DataTypesE.DLLReferences
               Return CheckFragmentsPositions(Text, ".:SF,.dll:EI")
            Case DataTypesE.EMailAddresses
               Return (CheckFragmentsPositions(Text, "@:M,.:M,.@:N,@.:N,..:N") AndAlso CheckCounts(Text, "@:1"))
            Case DataTypesE.GuidIds
               Return (CheckFragmentsPositions(Text, "{:S,}:E,-:M") AndAlso CheckCounts(Text, "-:4"))
            Case DataTypesE.HumanReadable
               Return True
            Case DataTypesE.Names
               Return CheckCounts(Text, "\,:1,(:1,):1, :1")
            Case DataTypesE.URLs
               Return CheckFragmentsPositions(Text, "\://:M")
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure displays a prompt and requests input from the user.
   Private Function GetInput(Prompt As String) As String
      Try
         Console.Write(Prompt)
         Return Console.ReadLine()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return ""
   End Function

   'This procedure displays a prompt and requests the user to select an option.
   Private Function GetSelection(Prompt As String, Selections As String, Optional DefaultSelection As String = Nothing) As String
      Try
         Dim KeyStroke As ConsoleKeyInfo = Nothing
         Dim Selection As String = DefaultSelection

         Console.Write(Prompt)
         Do
            KeyStroke = Console.ReadKey(intercept:=True)
            Select Case KeyStroke.Key
               Case ConsoleKey.Enter
                  If Not DefaultSelection = Nothing Then Exit Do
               Case ConsoleKey.Escape
                  Selection = Nothing
                  Exit Do
               Case Else
                  If Selections.Contains(KeyStroke.KeyChar) Then
                     Selection = KeyStroke.KeyChar
                     Exit Do
                  End If
            End Select
         Loop
         Console.WriteLine(Selection)

         Return Selection
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return ""
   End Function

   'This procedure handles any error that occur.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.ForegroundColor = ConsoleColor.Red
         Console.WriteLine()
         Console.WriteLine(ExceptionO.Message)
         Console.WriteLine($"Error code: { New Win32Exception(ExceptionO.Message).ErrorCode}")
         Console.WriteLine()
         Console.ForegroundColor = ConsoleColor.Gray
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim FileName As String = Nothing
         Dim IncludeUnicode As Boolean = False
         Dim SelectedDataType As String = CStr(DataTypesE.HumanReadable)

         My.Computer.FileSystem.CurrentDirectory = My.Application.Info.DirectoryPath

         Console.BackgroundColor = ConsoleColor.Black
         Console.ForegroundColor = ConsoleColor.Gray

         DisplayInformation()
         FileName = GetInput("Path: ")
         If FileName = Nothing Then Exit Sub
         If FileName.StartsWith("""") Then FileName = FileName.Substring(1)
         If FileName.EndsWith("""") Then FileName = FileName.Substring(0, FileName.Length - 1)

         Console.WriteLine()
         Console.WriteLine("0. DLL references")
         Console.WriteLine("1. E-Mail addresses")
         Console.WriteLine("2. GUID Ids")
         Console.WriteLine("3. Human readable (character codes 31-127)")
         Console.WriteLine("4. Names (last, initials (first))")
         Console.WriteLine("5. URLs")

         Console.WriteLine()
         SelectedDataType = GetSelection($"Data Type (default: {SelectedDataType}): ", "012345", "3")
         If SelectedDataType = Nothing Then Exit Sub

         IncludeUnicode = (GetSelection("Include unicode y/n? (default: y): ", "NYny", "y").ToLower() = "y")
         StartSearch(FileName, DirectCast(CInt(SelectedDataType), DataTypesE), IncludeUnicode)

         Console.BackgroundColor = ConsoleColor.Gray
         Console.ForegroundColor = ConsoleColor.Black
         Console.Write($" {SearchResultsList().Count} { If(SearchResultsList().Count = 1, "result", "results")} in ""{FileName}"" ")
         Console.BackgroundColor = ConsoleColor.Black
         Console.ForegroundColor = ConsoleColor.Gray
         Console.WriteLine()
         Console.WriteLine("Press any key to quit...")
         Do : Loop Until Console.KeyAvailable()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the search results list.
   Private Function SearchResultsList(Optional NewResult As String = Nothing, Optional ClearList As Boolean = False) As List(Of String)
      Try
         Static CurrentSearchResultsList As New List(Of String)

         If ClearList Then
            CurrentSearchResultsList.Clear()
         ElseIf Not NewResult = Nothing Then
            CurrentSearchResultsList.Add(NewResult)
         End If

         Return CurrentSearchResultsList
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of String)
   End Function

   'This procedure splits the text using the specified delimiter.
   Private Function SplitText(Text As String, Delimiter As Char, EscapeCharacter As Char) As List(Of String)
      Try
         Dim Position As Integer = Nothing
         Dim Texts As New List(Of String)

         If Not Text.EndsWith(Delimiter) Then Text &= Delimiter
         Do Until Text = Nothing
            Position = Text.IndexOf(Delimiter)
            If Position > 0 AndAlso Text.Chars(Position - 1) = EscapeCharacter Then
               Text = Text.Remove(Position - 1, 1)
               Position = Text.IndexOf(Delimiter, 1)
            End If
            Texts.Add(Text.Substring(0, Position))
            Text = Text.Substring(Position + 1)
         Loop

         Return Texts
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of String)
   End Function

   'This procedure starts the search for the specified data.
   Private Sub StartSearch(FileName As String, DataType As DataTypesE, IncludeUnicode As Boolean)
      Try
         SearchResultsList(, ClearList:=True)
         CurrentDataType(NewDataType:=DataType, NewIncludeUnicode:=IncludeUnicode)
         TextBin.FindText(TextBin.GetBinaryData(FileName))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure is called each time a string that fits the specified text definition is found.
   Private Sub TextBin_FoundText(Text As String, ByRef ContinueSearch As Boolean) Handles TextBin.FoundText
      Try
         Text = Text.Trim()

         If Not Text = Nothing Then
            If FitsDataType(Text) Then
               Select Case CurrentDataType()
                  Case DataTypesE.DLLReferences, DataTypesE.EMailAddresses, DataTypesE.URLs
                     Text = Text.ToLower()
                  Case DataTypesE.GuidIds
                     Text = Text.ToUpper()
               End Select

               If Not SearchResultsList().Contains(Text) Then
                  SearchResultsList(NewResult:=Text)
                  Console.WriteLine(Text)
               End If
            End If
         End If

         ContinueSearch = If(Console.KeyAvailable, Not (Console.ReadKey(intercept:=True).Key = ConsoleKey.Escape), True)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure is called when an error occurs in the Text Bin class.
   Private Sub TextBin_HandleError(TextBinExceptionO As Exception) Handles TextBin.HandleError
      Try
         HandleError(TextBinExceptionO)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Module
