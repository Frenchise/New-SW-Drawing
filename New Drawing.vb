Imports SwConst
'Imports SwCommandsSHEET_Template
Imports SldWorks
Imports System.IO

Public Class New_Drawing

	Dim swApp As SldWorks.SldWorks
	Dim swDraw As DrawingDoc
	Dim swDoc As ModelDoc2
	Dim CurDoc As ModelDoc2
	Dim swSheet As Sheet
	Dim CusProperties As CustomPropertyManager
	Dim CusProperties_Part As CustomPropertyManager
	Dim CusProperties_Assy As CustomPropertyManager
	Dim swModelDocExt As ModelDocExtension
	Dim swModelDocExt_Part As ModelDocExtension
	Dim swModelDocExt_Assy As ModelDocExtension
	Dim Bool_Result As Boolean
	ReadOnly DRAW_Template As String = "Drawing template"
	ReadOnly SHEET_Template As String = "Sheet 1 Template"
	ReadOnly ADD_SHEET_Template As String = "Additional sheet Template"
	Dim PC_USER As String = System.Environment.UserName
	Dim FullPathName As String
	Dim swView As View
	Dim annotations As Object
	Dim User_Saved As Boolean = False
	Dim Add_Sheet_Count As Integer = 3
	Dim Sheet_Count As Integer = 0
	Dim Drawing_Name_Saved As String
	Dim Add_Bom As Boolean = False
	Dim Clicked As Boolean = False
	Dim Title_ As String
	Dim Dir As String = "Default Directory"
	Dim Aircraft_Num As String
	Dim Add_View_Page As Integer
	Dim Opened_Files_Names As New List(Of String)

	Dim Part_Count As Integer = 0
	Dim Assy_Count As Integer = 0
	Dim Same_Page As Boolean = False

	'Dim Instance_list As New List(Of String)()
	Dim InstanceID_Added As Boolean = False
	Dim Main_Assy_Name As String


	Private Sub Form_Resize() Handles Me.ResizeEnd
		Functions.Form_resize(Me)
	End Sub

	Private Sub New_Drawing_Load(sender As Object, e As EventArgs) Handles MyBase.Load, Reload.Click

		Dim PC_USER As String = System.Environment.UserName
		Dim Start_Char As String


		SWFunctions.swAssy_Docs.Clear()
		SWFunctions.swPart_Docs.Clear()
		SWFunctions.swDwg_Docs.Clear()

		Title_ = Me.Text
		Me.Height = 250

		Form_Resize()
		
		'Initialize form to fit your needs
	

	End Sub




	Private Sub Generate_Drawing_Click(sender As Object, e As EventArgs) Handles Generate_Drawing.Click

		Dim count As Integer
		'Dim index As Integer
		Dim Open_Docs As Object

		If Clicked = False Then

			'Changes form back color to look busy"
			Me.BackColor = SystemColors.InactiveCaption
			Me.Text = Title_ + " - Loading"

			swApp = CreateObject("SldWorks.Application")
			swDoc = swApp.ActiveDoc

			count = swApp.GetDocumentCount
			Open_Docs = swApp.GetDocuments

				Dim swRootComp As Component2
				Dim swConfMgr As ConfigurationManager
				Dim swConf As Configuration

				swConfMgr = swDoc.ConfigurationManager
				swConf = swConfMgr.ActiveConfiguration
				swRootComp = swConf.GetRootComponent3(True) 'Top level Assy
				Main_Assy_Name = swRootComp.Name2

				If swDoc.GetType = swDocumentTypes_e.swDocASSEMBLY Then
					SWFunctions.Rename_Files = True
					SWFunctions.Add_Docs2(swRootComp, 1)
					SWFunctions.Rename_Files_with_PN()
					SWFunctions.Out_Put()
				End If

			Me.BackColor = SystemColors.Window
			Me.Text = Title_
			
		End If
	End Sub


	Private Sub Opened_ASSY_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Opened_ASSY.DoubleClick

		Dim count As Integer
		Dim index As Integer
		Dim Open_Docs As Object
		Dim Sheet_name As String
		Dim Sheet_Num As String
		Dim Last_Sheet As String
		Dim swSheetNames As Object
		Dim Same_Page As Boolean = False
		Dim Boundary_Box_Size As Double() = {0.0, 0.0, 0.0, 0.0}
		Dim Scale_Factor As Double
		Dim Prop_Name As String

		Dim i As Integer
		Dim status As Boolean
		Dim View_Count As Integer


		Dim sw_View As View
		Dim swAssy As AssemblyDoc

		Dim View_Name = String.Empty
		Dim First_View_Name = String.Empty
		Dim ValOut = String.Empty
		Dim ResolvedValOut = String.Empty
		Dim wasResolved As Boolean
		Dim linkToProp As Boolean

		Dim File_Name As String

		'Dim File_Name As String = Opened_ASSY.Text
		If Assy_Count = 0 Then
			File_Name = SWFunctions.swAssy_Docs(Assy_Count).Comp
		Else
			File_Name = SWFunctions.swAssy_Docs(Assy_Count).subcomp
		End If


		Drawing_Name_Saved = swDoc.GetPathName()

		swApp.ActivateDoc(File_Name)
		
		'Stole from API site
		Dim swModel As ModelDoc2
		Dim swConfMgr As ConfigurationManager
		Dim swConf As Configuration
		Dim swRootComp As Component2

		swModel = swApp.ActiveDoc
		swConfMgr = swModel.ConfigurationManager
		swConf = swConfMgr.ActiveConfiguration
		swRootComp = swConf.GetRootComponent3(True) 'Top level Assy
		Main_Assy_Name = swRootComp.Name2
		
		swAssy = swApp.ActiveDoc

		swModelDocExt_Assy = swAssy.Extension
		CusProperties_Assy = swModelDocExt_Assy.CustomPropertyManager("")
		View_Name = CusProperties_Assy.Get6("NOTE INFO", True, ValOut, ResolvedValOut, wasResolved, linkToProp)

		FullPathName = swDoc.GetPathName()

		'Checks if drawing document has been saved previously
		If User_Saved = False Then

			Open_New_Draw()

		Else

			Same_Assy_Page_Checkbox.Visible = True

			If Same_Assy_Page_Checkbox.Checked = True Then
				Same_Page = True
				Add_View_Page += 1
			Else
				Same_Page = False
				Add_View_Page = 0
			End If

			count = swApp.GetDocumentCount
			Open_Docs = swApp.GetDocuments

			For index = LBound(Open_Docs) To UBound(Open_Docs)
				swDoc = Open_Docs(index)
				FullPathName = swDoc.GetPathName()
				Name = SWFunctions.SW_FileName(FullPathName)

				If swDoc.GetType = 3 Then
					Name = SWFunctions.SW_FileName(FullPathName)
					swApp.ActivateDoc(Name)

				End If
			Next

		End If

		Add_Sheet_Count = swDraw.GetSheetCount
		swSheet = swDraw.GetCurrentSheet

		swSheetNames = swDraw.GetSheetNames
		Last_Sheet = swSheetNames(UBound(swSheetNames))

		If Same_Page = False Then

			If swSheet.GetName = Last_Sheet Then

				Add_Sheet_Count += 1
				Sheet_Num = Add_Sheet_Count.ToString
				Sheet_name = "Sheet" + Sheet_Num
			swDraw.NewSheet3(Sheet_name, 12, 12, 1, 1, False, ADD_SHEET_Template, 0, 0, "")
				swDraw.ActivateSheet(Sheet_name)

			Else
				swDraw.SheetNext()
			End If

		End If

		sw_View = swDraw.GetFirstView
		First_View_Name = sw_View.Name

		If ResolvedValOut <> "" Then

			Bool_Result = swDraw.Create3rdAngleViews2(File_Name)

			sw_View = sw_View.GetNextView
			If Same_Page = True Then

				View_Count = Add_View_Page * 3
				For i = 0 To View_Count - 1 '2
					sw_View = sw_View.GetNextView
				Next
				Dim View_Position As Double() = {0.4318, 0.1143}
				sw_View.Position = View_Position
			End If
			sw_View.SetName2(ResolvedValOut)
			Boundary_Box_Size = sw_View.GetOutline()

		Else
			Bool_Result = swDraw.Create3rdAngleViews2(File_Name)
		End If

		Dim sw_view1 As View = Nothing
		Dim sw_view2 As View = Nothing
		Dim sw_view3 As View = Nothing

		If Same_Page = False Then

			swDraw.ActivateView(ResolvedValOut)
			'Error gets thrown when the view name has been used already
			'can be on the same sheet or a different sheet
			'Todo: Check all views and names and rename the view if it's already used on a different sheet
			sw_view1 = swDraw.ActiveDrawingView
		End If
		If sw_view1 Is Nothing Then
			sw_view1 = swDraw.GetFirstView
			sw_view1 = sw_view1.GetNextView
			If Same_Page = True Then
				For j = 0 To View_Count - 1
					sw_view1 = sw_view1.GetNextView
				Next

			End If
		End If

		sw_view2 = sw_view1.GetNextView()
		If sw_view1.Name = ResolvedValOut Then
			sw_view2.SetName2(ResolvedValOut & " - TOP")
		End If

		sw_view3 = sw_view2.GetNextView()
		If sw_view1.Name = ResolvedValOut Then
			sw_view3.SetName2(ResolvedValOut & " - RIGHT")
		End If

		Scale_Factor = SWFunctions.View_Scale(sw_view1.GetOutline(), sw_view2.GetOutline(), sw_view3.GetOutline())

		sw_View.ScaleRatio() = {1, Scale_Factor}

		status = swDraw.ActivateView(ResolvedValOut)
		annotations = swDraw.InsertModelAnnotations3(0, 64, True, False, False, True)

		If Same_Page = False And User_Saved = False Then
			SWFunctions.Add_NoteInfo2(swDoc, ResolvedValOut, File_Name, SWFunctions.swAssy_Docs, SWFunctions.swPart_Docs, Add_BOM_Hardware.Checked)
			'SWFunctions.Add_NoteInfo2(swDoc, ResolvedValOut, File_Name, swFiles, Add_BOM_Hardware.Checked)
		End If

		If User_Saved = False Then

			SWFunctions.Save_Doc()

		End If

		Drawing_Name_Saved = swDoc.GetPathName()

		Dim Second_Warning As Integer = 1

		While Drawing_Name_Saved = "" And Second_Warning < 2

			User_Saved = False
			Functions.Error_Form("Save The File", "Please Save The Drawing File",,,, False, Start_Window)
			'''''''''''''''''''''''''''''''
			'SWFunctions.Save_Doc()
			Drawing_Name_Saved = swDoc.GetPathName()
			Second_Warning += 1

		End While

		If Drawing_Name_Saved <> "" Then
			User_Saved = True

		End If

		swDoc.ForceRebuild3(True)

		swApp.CloseDoc(File_Name)
		Same_Assy_Page_Checkbox.Checked = False

		Assy_Count += 1

	End Sub

	Private Sub Opened_Part_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Opened_Part.DoubleClick

		Dim count As Integer
		Dim index As Integer
		Dim Open_Docs As Object
		Dim Sheet_Name As String
		Dim Sheet_Num As String
		Dim Last_Sheet As String
		Dim swSheetNames As Object

		Dim Boundary_Box_Size As Double() = {0.0, 0.0, 0.0, 0.0}
		Dim Scale_Factor As Double
		Dim File_Extension As String
		Dim Prop_Name As String

		Dim sw_View As View
		Dim swPart As PartDoc
		'Dim swDoc As ModelDoc2
		'Dim swExtension As ModelDocExtension

		Dim View_Name = String.Empty
		Dim View_Count As Integer

		Dim ValOut = String.Empty
		Dim ResolvedValOut = String.Empty
		Dim wasResolved As Boolean
		Dim linkToProp As Boolean

		Dim File_Name As String = SWFunctions.swPart_Docs(Part_Count).subcomp

		Drawing_Name_Saved = swDoc.GetPathName()

		swApp.ActivateDoc(File_Name)

		swPart = swApp.ActiveDoc

		swModelDocExt_Part = swPart.Extension
		CusProperties_Part = swModelDocExt_Part.CustomPropertyManager("")
		View_Name = CusProperties_Part.Get6("NOTE INFO", True, ValOut, ResolvedValOut, wasResolved, linkToProp)

		FullPathName = swDoc.GetPathName()

		File_Extension = SWFunctions.SW_Extension(FullPathName)

		If Drawing_Name_Saved <> "" Then
			If File_Extension = ".SLDDRW" Then
				User_Saved = True
			ElseIf File_Extension <> ".SLDDRW" Then
				User_Saved = False
				Functions.Error_Form("Drawing not saved", "Please save the Part before continuing.",,,, False, Me)
				Exit Sub
			End If

		Else
			User_Saved = False
		End If

		If User_Saved = False Then
			Functions.Error_Form("Drawing not saved", "Please save the drawing before continuing.",,,, False, Me)
		Else

			'Open_New_Draw()

			Same_Parts_Page_Checkbox.Visible = True

			If Same_Page = True Then
				Add_View_Page += 1
			Else
				Add_View_Page = 0
			End If

			count = swApp.GetDocumentCount
			Open_Docs = swApp.GetDocuments

			For index = LBound(Open_Docs) To UBound(Open_Docs)
				swDoc = Open_Docs(index)
				FullPathName = swDoc.GetPathName()
				Name = SWFunctions.SW_FileName(FullPathName)

				If swDoc.GetType = 3 Then
					Name = SWFunctions.SW_FileName(FullPathName)
					swApp.ActivateDoc(Name)

				End If
			Next

			swDraw = swDoc 'Not sure why I did this

			swSheet = swDraw.GetCurrentSheet

			swSheetNames = swDraw.GetSheetNames
			Last_Sheet = swSheetNames(UBound(swSheetNames))

			If Same_Page = False Then

				If swSheet.GetName = Last_Sheet Then

					Add_Sheet_Count += 1
					Sheet_Num = Add_Sheet_Count.ToString
					Sheet_Name = "Sheet" + Sheet_Num
					swDraw.NewSheet3(Sheet_Name, 12, 12, 1, 1, False, ADD_SHEET_Template, 0, 0, "")
					swDraw.ActivateSheet(Sheet_Name)

				Else
					swDraw.SheetNext()

				End If

			End If


			sw_View = swDraw.GetFirstView

			If ResolvedValOut <> "" Then

				Bool_Result = swDraw.Create3rdAngleViews2(File_Name)

				sw_View = swDraw.GetFirstView
				sw_View = sw_View.GetNextView
				If Same_Page = True Then

					Dim SS As Object
					Dim VV As Object
					Dim SheetCount As Integer
					Dim Views As Integer

					SS = swDraw.GetViews
					For SheetCount = LBound(SS) To UBound(SS)

						VV = SS(SheetCount)

						For Views = LBound(VV) To UBound(VV)

						Next

					Next

					'View_Count = swDraw.GetViewCount - swDraw.GetSheetCount
					View_Count = Add_View_Page * 3
					'MsgBox(View_Count)
					For i = 0 To View_Count - 1
						sw_View = sw_View.GetNextView
					Next
					Dim View_Position As Double() = {0.4318, 0.1143}
					sw_View.Position = View_Position
				End If
				sw_View.SetName2(ResolvedValOut)
				Boundary_Box_Size = sw_View.GetOutline()

			Else
				Bool_Result = swDraw.Create3rdAngleViews2(File_Name)
			End If

			Dim sw_view1 As View = Nothing
			Dim sw_view2 As View = Nothing
			Dim sw_view3 As View = Nothing

			If Same_Page = False Then

				swDraw.ActivateView(ResolvedValOut)
				'Error gets thrown when the view name has been used already
				'can be on the same sheet or a different sheet
				'Todo: Check all views and names and rename the view if it's already used on a different sheet
				sw_view1 = swDraw.ActiveDrawingView
			End If
			If sw_view1 Is Nothing Then
				sw_view1 = swDraw.GetFirstView
				sw_view1 = sw_view1.GetNextView
				If Same_Page = True Then
					For j = 0 To View_Count - 1
						sw_view1 = sw_view1.GetNextView
					Next

				End If
			End If

			sw_view2 = sw_view1.GetNextView()
			If sw_view1.Name = ResolvedValOut Then
				sw_view2.SetName2(ResolvedValOut & " - TOP")
			End If

			sw_view3 = sw_view2.GetNextView()
			If sw_view1.Name = ResolvedValOut Then
				sw_view3.SetName2(ResolvedValOut & " - RIGHT")
			End If

			Scale_Factor = SWFunctions.View_Scale(sw_view1.GetOutline(), sw_view2.GetOutline(), sw_view3.GetOutline())

			sw_view1.ScaleRatio() = {1, Scale_Factor}

			annotations = swDraw.InsertModelAnnotations3(0, 32768, True, True, False, True) 'Inserts dimensions marked for drawing
			annotations = swDraw.InsertModelAnnotations3(0, 1048576, True, True, False, True) 'Inserts Hole Callout

			swDoc.ForceRebuild3(True)

			Part_Count += 1

		End If

		swApp.CloseDoc(File_Name)
		Same_Page = Not (Same_Page)

	End Sub

	Private Sub Open_New_Draw()

		Dim FolderName As String
		Dim swTitle As String
		Dim AC_Num As String
		Dim Drawings_Major As String
		Dim Folder_Check As String
		Dim Minor_Check As String
		Dim Structural As String
		Dim Note_Style As String = Nothing
		Dim Bool As Boolean

		Dim swAnnotation As Annotation
		Dim swNote As Note
		Dim swView As View
		Dim swLayerMgr As LayerMgr
		Dim Sheet1_Notes As Object

		swDoc = swApp.NewDocument(DRAW_Template, 0, 0, 0)
		swDraw = swDoc
		swModelDocExt = swDoc.Extension
		CusProperties = swModelDocExt.CustomPropertyManager("")

		If Request.Text.ToString = "" Then
			Request.Text = "N/A"
		Else
			CusProperties.Set2("RELEASED TO", Request.Text.ToString.ToUpper)
			CusProperties.Set2("REQUESTED BY", Request.Text.ToString.ToUpper)

		End If

		Select Case PC_USER

		End Select

		CusProperties.Set2("DATE", DateTime.Now.ToString("MMM/dd/yyyy").ToUpper)
		CusProperties.Set2("DRAWN BY", PC_USER)
		CusProperties.Set2("CURRENT REV", "$PRP:""REV 1""")
		CusProperties.Set2("REV DESCRIPTION", "PRELIMINARY RELEASE")

		'Adding Notes to Sheet 1
		swDraw.ActivateSheet("Sheet1")

		swLayerMgr = swDoc.GetLayerManager

		If Note_Style IsNot Nothing Then

			swView = swDraw.GetFirstView
			Sheet1_Notes = swView.GetNotes

			For i = 0 To UBound(Sheet1_Notes)

				swNote = Sheet1_Notes(i)
				If swNote.GetName = "Drawing-Notes" Then

					swAnnotation = swNote.GetAnnotation

					swAnnotation.SetPosition2(1.15 / 39.3701, 9.75 / 39.3701, 0)
					swAnnotation.SetStyleName(Note_Style)
					swAnnotation.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
					swAnnotation.Layer = "NOTES"

				End If
			Next
		End If
		Bool = swLayerMgr.SetCurrentLayer("-Per Standard-")
	End Sub

	Private Sub Add_All_Click(sender As Object, e As EventArgs) Handles Add_All.Click

		Generate_Drawing.PerformClick()

		For i = 0 To SWFunctions.swAssy_Docs.Count - 1 'Opened_ASSY.Items.Count - 1
			'Opened_ASSY.SetSelected(i, True)
			Opened_ASSY_SelectedIndexChanged(sender, e)
			'SWFunctions.Add_Docs(swRootComp, 1)

		Next

		For j = 0 To SWFunctions.swPart_Docs.Count - 1 'Opened_Part.Items.Count - 1

			'Opened_Part.SetSelected(j, True)
			Opened_Part_SelectedIndexChanged(sender, e)
		Next

		Functions.Error_Form("Done", "All Assemblies and Part files have been added",,,, False,)
		SWFunctions.Out_Put()
		Me.Close()
	End Sub



End Class
