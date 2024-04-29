Option Explicit
Option Compare Text

Dim msg As String


Sub ProcessJobs()

    msg = "Processing messages in Job Market folder..."
    MsgBox (msg)
   
    Dim ns_mapi As Outlook.NameSpace
    Dim d_folder As Outlook.Folder        ' default folder
    Dim j_folder As Outlook.Folder        ' jobs folder
    
    Set ns_mapi = Outlook.Application.GetNamespace("MAPI")
    Set d_folder = ns_mapi.GetDefaultFolder(olFolderInbox)
    Set j_folder = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("Job Market")
    
    
    Dim j_folder2 As Outlook.Folder        ' jobs folder
    Dim j_items As Outlook.Items          ' jobs items
    Dim j_item As Outlook.MailItem        ' job item (mail message)
    Dim n_item As Outlook.MailItem        ' new item ( copy job item to destination folder
    
    Dim skillKeys As Variant
    Dim placeKeys As Variant
    
    Dim CAFolder As Outlook.Folder
    Dim PAFolder As Outlook.Folder
    Dim NJFolder As Outlook.Folder
    Dim OTFolder As Outlook.Folder
    
    Dim crp_count As Integer
    Dim job_count As Integer
    
    skillKeys = Array("java", "python", ".net", "fullstack")
    placeKeys = Array("CA", "PA", "NJ")
    
    Let crp_count = 0
    Let job_count = 0
    
    'Source
    Set j_items = j_folder.Items
    
    'Targets
    Set CAFolder = j_folder.Folders("CA Jobs")
    Set PAFolder = j_folder.Folders("PA Jobs")
    Set NJFolder = j_folder.Folders("NJ Jobs")
    Set OTFolder = j_folder.Folders("Other Jobs")
    
    Dim strCategory As String
    Dim interested As Boolean
    Dim i As Integer
    
    'On Error GoTo ErrorHandler
    On Error Resume Next
    For Each j_item In j_items
    
        interested = False
        strCategory = ""
        
        If InStr(LCase(j_item.Subject), "java") Then
            If strCategory > "" Then strCategory = strCategory + ","
            strCategory = strCategory + "~Java"
            interested = True
        End If
        
        If InStr(LCase(j_item.Subject), ".net") Then
            If strCategory > "" Then strCategory = strCategory + ","
            strCategory = "~.NET"
            interested = True
        End If
        
        If InStr(LCase(j_item.Subject), "python") Then
            If strCategory > "" Then strCategory = strCategory + ","
            strCategory = "~Python"
            interested = True
        End If
        
        If InStr(LCase(j_item.Subject), "fullstack") Then
            If strCategory > "" Then strCategory = strCategory + ","
            strCategory = "~Full Stack"
            interested = True
        End If
        
        j_item.Categories = strCategory
        
        If interested Then
            
            interested = False
            
            If InStr(j_item.Subject, "CA") Then
                Set n_item = j_item.Copy
                n_item.Move CAFolder
                interested = True
            End If
            
            If InStr(j_item.Subject, "PA") Then
                Set n_item = j_item.Copy
                n_item.Move PAFolder
                interested = True
            End If
            
            If InStr(j_item.Subject, "NJ") Then
                Set n_item = j_item.Copy
                n_item.Move NJFolder
                interested = True
            End If
            
            If interested Then
                job_count = job_count + 1
                j_item.Delete
            Else
                crp_count = crp_count + 1
                j_item.Move OTFolder
            End If
            
            
            
        End If
        
    Next   'getting a type mismatch here.
    
    Set j_item = Nothing
    Set j_items = Nothing
    Set j_folder = Nothing
    Set d_folder = Nothing
    
    MsgBox (job_count & " : jobs. " & crp_count & " : crap")
    
    Set ns_mapi = Nothing
    
ProgramExit:
    Exit Sub
    
ErrorHandler:
    
    
End Sub




