Attribute VB_Name = "MainMenu_Emails"
Option Explicit
Public UnivFolderList As Variant
Public UnivUserName As String



Sub UserForm_Launcher()
    ' This function loads the MainMenu and shows it
    Load UF_MainMenu
    UF_MainMenu.Show
End Sub


'Outlook VB Macro to move selected mail item(s) to a target folder
Sub MoveToFolder(ByVal FolderList As Variant, Optional Inbox_UserName As String = "xxx")
    On Error Resume Next
    
    Dim ns As Outlook.NameSpace
    Dim MoveToFolder As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    
    
    Set ns = Application.GetNamespace("MAPI")
    
    Dim Inbox As Outlook.MAPIFolder
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    

    Dim objFolder As MAPIFolder
    Dim InboxName As String
    Dim objOwner As Outlook.Recipient
    Dim item As Variant
    
    If Not Inbox_UserName = "xxx" Then
        Set objOwner = ns.CreateRecipient(Inbox_UserName)
        objOwner.Resolve
        If objOwner.Resolved Then
            Set objFolder = ns.GetSharedDefaultFolder(objOwner, olFolderInbox)
        End If
        
        For Each item In FolderList
            Set objFolder = objFolder.Folders(item)
        Next item

    Else
        Set objFolder = ns.Folders(FolderList(0))
        Dim ArrayPos As Integer
        ArrayPos = 1
        While (ArrayPos <= UBound(FolderList))
            Set objFolder = objFolder.Folders(FolderList(ArrayPos))
            ArrayPos = ArrayPos + 1
        Wend
    End If
       
    Set MoveToFolder = objFolder
        
    

    
    If Application.ActiveExplorer.Selection.Count = 0 Then
       MsgBox ("No item selected")
       Exit Sub
    End If
    
    If MoveToFolder Is Nothing Then
       MsgBox "Target folder not found!", vbOKOnly + vbExclamation, "Move Macro Error"
    End If
    
    For Each objItem In Application.ActiveExplorer.Selection
       If MoveToFolder.DefaultItemType = olMailItem Then
          If objItem.Class = olMail Then
             objItem.Move MoveToFolder
          End If
      End If
    Next
    
    Set objItem = Nothing
    Set MoveToFolder = Nothing
    Set ns = Nothing
    
    
    If UF_MainMenu.Visible = True Then
        Unload UF_MainMenu
    End If
    
    If UF_FolderMenu.Visible = True Then
        Unload UF_FolderMenu
    End If

End Sub

Function Load_UF_FolderMenu(ByVal FolderList As Variant, Optional Inbox_UserName As String = "xxx") As Boolean
    Load_UF_FolderMenu = True
    Dim DictFolders As Scripting.Dictionary 'Contiene la información sobre la cantidad de coincidencias por carpeta
        Set DictFolders = New Scripting.Dictionary
        Dim x As Integer
        Dim MaxElements As Integer
        
        'If the folders are inside the Inbox of the given, search this way
        
        Set DictFolders = BuscarCarpetas(FolderList, Inbox_UserName)
    
        If DictFolders.Count < 8 Then
            MaxElements = DictFolders.Count
        Else
            MaxElements = 8
        End If
        
        If MaxElements = 0 Then
            MsgBox ("No matching element found")
            Load_UF_FolderMenu = False
        Else
            For x = 0 To MaxElements - 1
                'MsgBox (DictFolders.Keys()(x) & DictFolders.Items()(x))
                UF_FolderMenu.Controls("Label" & x + 1).Caption = DictFolders.Keys()(x)
            Next x
        End If
       
       
       For x = MaxElements + 1 To 8
            UF_FolderMenu.Controls("Label" & x).Visible = False
            UF_FolderMenu.Controls("Label" & x + 8).Visible = False
       
       Next x
End Function


Function BuscarCarpetas(ByVal FolderList As Variant, Optional Inbox_UserName As String = "xxx") As Object
    'Input MailBoxName As String 'Nombre del mailbox, normalmente Baul
    'Input SubFolder1Name As String
    'Input SubFolder2Name As String
    
    Dim StringToSearch As String 'String the user inputs, like the hashtag
    
    Dim FolderString As String 'Folder name which will be looked up
    Dim FolderString_original As String 'Found folder name (case sensitive)
    
    Dim SearchedWords As Object 'set of words in the input to be searched
    Set SearchedWords = CreateObject("System.Collections.ArrayList")
    
    Dim FolderWords As Object 'set of words of the iterated folder
    Set FolderWords = CreateObject("System.Collections.ArrayList")
    
    Dim DictFolderWords As Scripting.Dictionary 'tells how many matches there are in each folder
    Set DictFolderWords = New Scripting.Dictionary
    
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
   
    'Set of variables needed to move
    Dim IterateFolder As Outlook.MAPIFolder
    Dim IteratedSubFolder As Outlook.Folder
    Dim objItem As Outlook.MailItem
    
    'Integers for final iteration
    Dim WordInFolder As Integer
    Dim WordInString As Integer
    Dim x As Integer
    
    'Variables to get the sorted dictionary
    Set BuscarCarpetas = New Scripting.Dictionary
    
    Dim objFolder As MAPIFolder
    Dim item As Variant
    
    If Not Inbox_UserName = "xxx" Then
        Dim objOwner As Outlook.Recipient
        Set objOwner = ns.CreateRecipient(Inbox_UserName)
        objOwner.Resolve
        If objOwner.Resolved Then
            Set objFolder = ns.GetSharedDefaultFolder(objOwner, olFolderInbox)
        End If
        
        For Each item In FolderList
            Set objFolder = objFolder.Folders(item)
        Next item
    Else
        Set objFolder = ns.Folders(FolderList(0))
        Dim ArrayPos As Integer
        ArrayPos = 1
        While (ArrayPos <= UBound(FolderList))
            Set objFolder = objFolder.Folders(FolderList(ArrayPos))
            ArrayPos = ArrayPos + 1
        Wend
    End If
    
    
    StringToSearch = InputBox("Introduce hashtags to be searched", "Input mask")
    
    'String transformed to lower case
    StringToSearch = LCase(StringToSearch)
    
    If StrPtr(StringToSearch) = 0 Then
        StringToSearch = "agasgasdgasgasgasdga"
    End If
    Set SearchedWords = StringToArray(StringToSearch)
    
    For Each IteratedSubFolder In objFolder.Folders
        
        'Get Folder name
        FolderString_original = IteratedSubFolder.name
        
        'set to lower case
        FolderString = LCase(FolderString_original)
        
        'split and convert to array
        Set FolderWords = StringToArray(FolderString)
        
        For WordInFolder = 0 To FolderWords.Count - 1
            
            For WordInString = 0 To SearchedWords.Count - 1
                If FolderWords(WordInFolder) = SearchedWords(WordInString) Then
                    'Count ocurrencies
                    DictFolderWords(FolderString_original) = DictFolderWords(FolderString_original) + 1
                End If
            Next
        Next
               
    Next
    If DictFolderWords.Count = 0 Then
        Set BuscarCarpetas = New Scripting.Dictionary
    Else
        Set BuscarCarpetas = SortDicDescendiente(DictFolderWords)
    End If
    
End Function


Function StringToArray(ByVal s As String) As Object

    Dim WordsList As Object
    Set WordsList = CreateObject("System.Collections.ArrayList")
    Set StringToArray = CreateObject("System.Collections.ArrayList")
    Dim StrPos As Integer 'iterated position
    Dim PosIni As Integer 'initial position of the analyzed word
    Dim PosFin As Integer 'initial position of the analyzed word
    Dim Word As String
    
    PosIni = 1 'take first word
    
    For StrPos = 1 To Len(s)
        If (Mid(s, StrPos, 1) = " ") Or StrPos = Len(s) Then
            PosFin = StrPos - 1
            If StrPos = Len(s) Then
                PosFin = PosFin + 1
            End If
            Word = Mid(s, PosIni, PosFin - PosIni + 1)
            WordsList.Add Word
            PosIni = PosFin + 2
        End If
    Next
    
    Set StringToArray = WordsList.Clone
End Function

Public Function SortDicDescendiente(ByRef DictInput As Scripting.Dictionary, Optional ByVal Descendiente As Boolean = True) As Scripting.Dictionary
    'Sort dictionary in a descending order
    Dim DictSorted As Scripting.Dictionary
    Set DictSorted = New Scripting.Dictionary
    
    'Creamos un array para los nombres de los keys y otro para sus valores
    Dim listKey As Object
    Set listKey = CreateObject("System.Collections.ArrayList")
    Dim listValue As Object
    Set listValue = CreateObject("System.Collections.ArrayList")


    Dim DictPosition As Integer
    Dim ListLength As Integer
    Dim ListPosition As Integer
    

    listKey.Add DictInput.Keys(0)
    listValue.Add DictInput.Items(0)
    ListLength = 1


    For DictPosition = 1 To DictInput.Count - 1
        ListPosition = 0
        
        Do While True
            
            If DictInput.Items(DictPosition) > listValue(ListPosition) Then
                ListLength = ListLength + 1
                listKey.Insert ListPosition, DictInput.Keys(DictPosition)
                listValue.Insert ListPosition, DictInput.Items(DictPosition)
                GoTo Break
            End If
            
            If ListPosition = ListLength - 1 Then
                'Entonces es que es el último elemento y hay que insertarlo al final.
                ListLength = ListLength + 1
                listKey.Insert ListPosition + 1, DictInput.Keys(DictPosition)
                listValue.Insert ListPosition + 1, DictInput.Items(DictPosition)
                GoTo Break
            
            End If
            
            ListPosition = ListPosition + 1
        Loop
        
Break:
    Next DictPosition

    'Según si el input es descendiente=True o descendiente = False, añadiremos las variables al diccionario dictSorted de forma descendiente o ascendente
    If Descendiente Then
        For DictPosition = 0 To DictInput.Count - 1
            DictSorted.Add listKey(DictPosition), listValue(DictPosition)
        Next DictPosition
    Else
        For DictPosition = DictInput.Count - 1 To 0 Step -1
            DictSorted.Add listKey(DictPosition), listValue(DictPosition)
        Next DictPosition
    End If
    
    Set SortDicDescendiente = DictSorted
    
End Function

Sub SuggestFolders(ByVal FolderList As Variant, Optional Inbox_UserName As String = "xxx")
' This function is called from the UserForm of the Folders when clicking on advanced search (with hashtags)
' It unloads the previously shown userform and loads the folder menu
' Input is the list of folders in which it has to look
    UnivUserName = Inbox_UserName
    UnivFolderList = FolderList
    
    Dim Show As Boolean
    Load UF_FolderMenu
    Unload UF_MainMenu
    Show = Load_UF_FolderMenu(FolderList, Inbox_UserName)
    
    If Show Then
        UF_FolderMenu.Show
    Else
        Unload UF_FolderMenu
    End If
    

    
End Sub




