Imports SldWorks
Imports SwCommands
Imports SwConst
Imports System.IO
Imports System.IO.Compression
Imports System.IO.Compression.ZipFile
Imports Microsoft.Office.Interop.Excel

Public Class SWFunctions

	Public Shared swApp As SldWorks.SldWorks

	Public Shared CusProperties_Part As CustomPropertyManager
	Public Shared CusProperties_Assy As CustomPropertyManager
	Public Shared swModelDocExt_Part As ModelDocExtension
	Public Shared swModelDocExt_Assy As ModelDocExtension

	Public Shared swAssy_Docs As New List(Of Assy_Docs)
	Public Shared swPart_Docs As New List(Of Part_Docs)
	Public Shared swDwg_Docs As New List(Of Drawing_Docs)
	
	Shared swComp_Assy As String

	Class Assy_Docs
		Public Comp As String
		Public subcomp As String
		Public instance_ID As String
		Public Part_Number As String = "Null"
		Public Nomenclature As String = "Null"
		Public Spec As String = "Null"
		Public Description As String = "Null"
		Public Material As String = "Null"
		Public Weight As String = "Null"
		Public Parent As String = "Null"
		Public Name As String
		Public Counter As Integer
		Public Used As Boolean

		Public Sub New(s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, s6 As String, 
						s7 As String, s8 As String, s9 As String, Optional s10 As String = "")
			Comp = s1
			subcomp = s2
			instance_ID = s3
			Part_Number = s4
			Nomenclature = s5
			Description = s6
			Spec = s7
			Material = s8
			Weight = s9
			Parent = s10

		End Sub

	End Class

	Class Part_Docs
		Public Comp As String
		Public subcomp As String
		Public instance_ID As String
		Public Part_Number As String = "Null"
		Public Nomenclature As String = "Null"
		Public Spec As String = "Null"
		Public Description As String = "Null"
		Public Material As String = "Null"
		Public Weight As String = "Null"
		Public Parent As String = "Null"
		Public Name As String
		Public Counter As Integer
		Public Used As Boolean

		Public Sub New(s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, s6 As String, 
						s7 As String, s8 As String, s9 As String, Optional s10 As String = "")
			Comp = s1
			subcomp = s2
			instance_ID = s3
			Part_Number = s4
			Nomenclature = s5
			Description = s6
			Spec = s7
			Material = s8
			Weight = s9
			Parent = s10

		End Sub

	End Class

	
	
	Shared Function SW_FileName(path As String)
			Dim FileName As String = path

			FileName = FileName.Remove(0, FileName.LastIndexOf("\") + 1)
			FileName = FileName.Remove(FileName.LastIndexOf("."))

			Return FileName
		End Function


	Shared Function SW_Extension(path As String)
			Dim Extension As String = path

			Extension = Extension.Remove(0, Extension.LastIndexOf("."))

			Return Extension
		End Function


	
	
	Shared Function Save_Doc()
			Dim swApp As SldWorks.SldWorks
			Dim bool As Boolean
			Dim swDoc As ModelDoc2

			swApp = CreateObject("SldWorks.Application")
			swDoc = swApp.ActiveDoc

			bool = swDoc.Extension.RunCommand(swCommands_e.swCommands_SaveAs, "")
			Return bool
		End Function


	Shared Function View_Scale(Outline1() As Double, Optional Outline2() As Double = Nothing, Optional Outline3() As Double = Nothing)

			Dim View1 As Double = 10
			Dim View2 As Double = 10
			Dim View3 As Double = 10
			Dim Small_Scale As Double = 10 'Set as high value

			View1 = Boundary_box(Outline1)

			If Outline2 IsNot Nothing Then
				View2 = Boundary_box(Outline2)
			End If

			If Outline3 IsNot Nothing Then
				View3 = Boundary_box(Outline3)
			End If

			If View1 <> 1 Or View2 <> 1 Or View3 <> 1 Then
				Dim Scales As Double() = {View1, View2, View3}

				For Each element As Double In Scales
					Small_Scale = Math.Min(Small_Scale, element)
				Next

			End If

			Small_Scale = Math.Round(1 / Small_Scale, 1)


			Return Small_Scale
		End Function


	Shared Function Boundary_box(Outline() As Double)

			Dim View_Scale_Factor As Decimal() = {0.0, 0.0}
			Dim Boundary As Double() = {0.0, 0.0, 0.0, 0.0}
			Dim Scales_sw As Double() = {1 / 2, 1 / 3, 1 / 4, 1 / 5, 1 / 6, 1 / 7, 1 / 8, 1 / 10, 1 / 12, 1, 2 / 3, 2, 3}
			Dim Scale_Factor As Double = 0.0

			Boundary = Outline

			Boundary(0) = Boundary(0) * 39.3701 ' x min
			Boundary(1) = Boundary(1) * 39.3701 ' y min
			Boundary(2) = Boundary(2) * 39.3701 ' x max
			Boundary(3) = Boundary(3) * 39.3701 ' y max

			Boundary(2) = Boundary(2) - Boundary(0)
			Boundary(3) = Boundary(3) - Boundary(1)

			View_Scale_Factor(0) = 4.875 / Boundary(3) 'Scale factor to achieve 4.875" on Y height
			View_Scale_Factor(1) = 10.5 / Boundary(2) 'Scale factor to achieve 10.5" on X width

			Dim smallestdiff As Double = Math.Abs(View_Scale_Factor(0) - Scales_sw(0))
			Dim smallestdiffIndex = 0
			Dim Currdiff As Double
			Dim j = 0
			For j = 0 To Scales_sw.Count - 1
				Currdiff = Math.Abs(View_Scale_Factor(0) - Scales_sw(j))
				If Currdiff < smallestdiff Then
					smallestdiff = Currdiff
					smallestdiffIndex = j
				End If

			Next

			View_Scale_Factor(0) = Scales_sw(smallestdiffIndex)

			For j = 0 To Scales_sw.Count - 1
				Currdiff = Math.Abs(View_Scale_Factor(1) - Scales_sw(j))
				If Currdiff < smallestdiff Then
					smallestdiff = Currdiff
					smallestdiffIndex = j
				End If

			Next

			View_Scale_Factor(1) = Scales_sw(smallestdiffIndex)

			'Todo: Compare each View_Scale_Factor to common scale options, pick the one closest

			Scale_Factor = Math.Min(View_Scale_Factor(0), View_Scale_Factor(1))

			Return Scale_Factor
		End Function


	

'Add notes with linked properties tied to specific views and properties.
'Not as useful as it seems
'Duplicate Instance_Num counts to accomodate the late change to adding two different lists
	Shared Function Add_NoteInfo2(swDoc As ModelDoc2, Drawing_View As String,
								 File_name As String, Instance_Num_Assy As List(Of SWFunctions.Assy_Docs), 
								Instance_Num_Part As List(Of SWFunctions.Part_Docs), Optional ByVal Add_Bom As Boolean = False)
		

		Dim sw_View As View
		Dim swNote_Info As Note
		Dim swAnno_Info As Annotation
		Dim swNote_DESCRIPTION As Note = Nothing
		Dim swAnno_DESCRIPTION As Annotation
		Dim swNote_NOMENCLATURE As Note = Nothing
		Dim swAnno_NOMENCLATURE As Annotation
		Dim swNote_SPEC As Note = Nothing
		Dim swAnno_SPEC As Annotation
		Dim swLayerMgr As LayerMgr


		Dim Text_Add As String
		Dim NOTE_INFO As String
		Dim DESCRIPTION As String
		Dim NOMENCLATURE As String
		Dim SPEC As String
		Dim Bool As Integer

		Dim x_pos As Double = 0.0
		Dim x_pos_Nom As Double = 0.1
		Dim x_pos_Des As Double = 0.2
		Dim x_pos_Spec As Double = 0.3
		Dim y_pos As Double = 0
		Dim z_pos As Double = 0.0

		Dim ChartoTrim As Char() = {"@"}

		swLayerMgr = swDoc.GetLayerManager
		Bool = swLayerMgr.SetCurrentLayer("NOTES")

		sw_View = swDoc.GetFirstView
		sw_View = sw_View.GetNextView

		NOTE_INFO = "$PRPSMODEL:" & Chr(34) & "NOTE INFO" & Chr(34) & " $COMP:" & Chr(34) & File_name & "@" & Drawing_View & "/"

		DESCRIPTION = "$PRPSMODEL:" & Chr(34) & "DESCRIPTION" & Chr(34) & " $COMP:" & Chr(34) & File_name & "@" & Drawing_View & "/"
		NOMENCLATURE = "$PRPSMODEL:" & Chr(34) & "NOMENCLATURE" & Chr(34) & " $COMP:" & Chr(34) & File_name & "@" & Drawing_View & "/"
		SPEC = "$PRPSMODEL:" & Chr(34) & "SPEC" & Chr(34) & " $COMP:" & Chr(34) & File_name & "@" & Drawing_View & "/"

		For i = 0 To Instance_Num_Assy.Count  'Files - 1

			If i = 0 Then

				Text_Add = "$PRPSMODEL:" & Chr(34) & "NOTE INFO" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" & Drawing_View & Chr(34)
				swNote_Info = swDoc.InsertNote(Text_Add)

				If New_Drawing.Add_BOM_Hardware.Checked = True Then
					Text_Add = "$PRPSMODEL:" & Chr(34) & "DESCRIPTION" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" 
											& Drawing_View & Chr(34)
					swNote_DESCRIPTION = swDoc.InsertNote(Text_Add)

					Text_Add = "$PRPSMODEL:" & Chr(34) & "NOMENCLATURE" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" 
											& Drawing_View & Chr(34)
					swNote_NOMENCLATURE = swDoc.InsertNote(Text_Add)

					Text_Add = "$PRPSMODEL:" & Chr(34) & "SPEC" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" 
											& Drawing_View & Chr(34)
					swNote_SPEC = swDoc.InsertNote(Text_Add)
				End If
			Else

				Text_Add = NOTE_INFO + Instance_Num_Assy(i - 1).instance_ID & Chr(34)
				swNote_Info = swDoc.InsertNote(Text_Add)

				If Add_Bom = True Then

					Text_Add = DESCRIPTION + Instance_Num_Assy(i - 1).instance_ID & Chr(34)
					swNote_DESCRIPTION = swDoc.InsertNote(Text_Add)

					Text_Add = NOMENCLATURE + Instance_Num_Assy(i - 1).instance_ID & Chr(34)
					swNote_NOMENCLATURE = swDoc.InsertNote(Text_Add)

					Text_Add = SPEC + Instance_Num_Assy(i - 1).instance_ID & Chr(34)
					swNote_SPEC = swDoc.InsertNote(Text_Add)
				End If
			End If

			swAnno_Info = swNote_Info.GetAnnotation()
			swAnno_Info.SetAttachedEntities(sw_View)
			swAnno_Info.SetPosition2(x_pos, y_pos, z_pos)

			If Add_Bom = True And i <> 0 Then
				swAnno_NOMENCLATURE = swNote_NOMENCLATURE.GetAnnotation()
				swAnno_NOMENCLATURE.SetAttachedEntities(sw_View)
				swAnno_NOMENCLATURE.SetPosition2(x_pos_Nom, y_pos, z_pos)

				swAnno_DESCRIPTION = swNote_DESCRIPTION.GetAnnotation()
				swAnno_DESCRIPTION.SetAttachedEntities(sw_View)
				swAnno_DESCRIPTION.SetPosition2(x_pos_Des, y_pos, z_pos)

				swAnno_SPEC = swNote_SPEC.GetAnnotation()
				swAnno_SPEC.SetAttachedEntities(sw_View)
				swAnno_SPEC.SetPosition2(x_pos_Spec, y_pos, z_pos)
			End If


			y_pos -= 0.00889

		Next

		For i = 0 To Instance_Num_Part.Count  'Files - 1

			If i = 0 Then

				Text_Add = "$PRPSMODEL:" & Chr(34) & "NOTE INFO" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" & Drawing_View & Chr(34)
				swNote_Info = swDoc.InsertNote(Text_Add)

				If New_Drawing.Add_BOM_Hardware.Checked = True Then
					Text_Add = "$PRPSMODEL:" & Chr(34) & "DESCRIPTION" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" 
											& Drawing_View & Chr(34)
					swNote_DESCRIPTION = swDoc.InsertNote(Text_Add)

					Text_Add = "$PRPSMODEL:" & Chr(34) & "NOMENCLATURE" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" 
											& Drawing_View & Chr(34)
					swNote_NOMENCLATURE = swDoc.InsertNote(Text_Add)

					Text_Add = "$PRPSMODEL:" & Chr(34) & "SPEC" & Chr(34) & " $COMP:" & Chr(34) & File_name & "-1" & "@" 
											& Drawing_View & Chr(34)
					swNote_SPEC = swDoc.InsertNote(Text_Add)
				End If
			Else

				Text_Add = NOTE_INFO + Instance_Num_Part(i - 1).instance_ID & Chr(34)

				swNote_Info = swDoc.InsertNote(Text_Add)

				If Add_Bom = True Then

					Text_Add = DESCRIPTION + Instance_Num_Part(i - 1).instance_ID & Chr(34)
					swNote_DESCRIPTION = swDoc.InsertNote(Text_Add)

					Text_Add = NOMENCLATURE + Instance_Num_Part(i - 1).instance_ID & Chr(34)
					swNote_NOMENCLATURE = swDoc.InsertNote(Text_Add)

					Text_Add = SPEC + Instance_Num_Part(i - 1).instance_ID & Chr(34)
					swNote_SPEC = swDoc.InsertNote(Text_Add)
				End If
			End If

			swAnno_Info = swNote_Info.GetAnnotation()
			swAnno_Info.SetAttachedEntities(sw_View)
			swAnno_Info.SetPosition2(x_pos, y_pos, z_pos)

			If Add_Bom = True And i <> 0 Then
				swAnno_NOMENCLATURE = swNote_NOMENCLATURE.GetAnnotation()
				swAnno_NOMENCLATURE.SetAttachedEntities(sw_View)
				swAnno_NOMENCLATURE.SetPosition2(x_pos_Nom, y_pos, z_pos)

				swAnno_DESCRIPTION = swNote_DESCRIPTION.GetAnnotation()
				swAnno_DESCRIPTION.SetAttachedEntities(sw_View)
				swAnno_DESCRIPTION.SetPosition2(x_pos_Des, y_pos, z_pos)

				swAnno_SPEC = swNote_SPEC.GetAnnotation()
				swAnno_SPEC.SetAttachedEntities(sw_View)
				swAnno_SPEC.SetPosition2(x_pos_Spec, y_pos, z_pos)
			End If


			y_pos -= 0.00889

		Next

		Return True

	End Function

'Heavily modified GetChildren method (IComponent2) to fit my specific needs. Not clean or intuitive to understand by reading.
'This was the hardest part for me to get a specific outcome to add documents.
'Lots of trial and error with unforeseen outcomes
	Shared Function Add_Docs2(ByVal swComp As Component2, ByVal nLevel As Integer)

		swApp = CreateObject("SldWorks.Application")

		Dim swPart As ModelDoc2
		Dim swPart2 As ModelDoc2
		Dim swPart3 As ModelDoc2
		Dim swChildComp As Component2
		Dim swParent As Component2

		Dim vChildComp As Object

		Dim Status As Boolean = False
		Dim Used As Boolean = False
		Dim Add_To_List As Boolean = False
		Dim isAssy As Boolean = False
		Dim isPart As Boolean = False
		Dim Assy_Count As Integer


		Dim ValOut = String.Empty
		Dim wasResolved As Boolean
		Dim linkToProp As Boolean
		Dim Dash_Name = String.Empty
		Dim Temp_Name = String.Empty
		Dim Parent_name = String.Empty

		Dim errorval As Integer

		Dim CusProp As String() = {"PART NUMBER", "NOMENCLATURE", "DESCRIPTION", "SPEC", "MATERIAL", "WEIGHT"}
		Dim RecAssyCusProp As String() = {"N/A", "N/A", "N/A", "N/A", "N/A", "N/A"}


		If nLevel = 1 Then

			swComp_Assy = swComp.Name2
			swPart2 = swComp.GetModelDoc2

			If swPart2.GetType = swDocumentTypes_e.swDocASSEMBLY Then

				swModelDocExt_Assy = swPart2.Extension
				CusProperties_Assy = swModelDocExt_Assy.CustomPropertyManager("")

				For propNum = 0 To UBound(CusProp)
					Dash_Name = CusProperties_Assy.Get6(CusProp(propNum), True, ValOut, RecAssyCusProp(propNum), wasResolved, linkToProp)
				Next

				If CusProperties_Assy.Get6(CusProp(0), True, ValOut, RecAssyCusProp(0), wasResolved, linkToProp) = 2 Then
					If RecAssyCusProp(0) <> "" And RecAssyCusProp(0) <> "-XX" Then

						Temp_Name = RecAssyCusProp(0).Substring(0, 1)

						If Temp_Name = "-" Then
							Add_To_List = True
						End If
					End If

				End If
				If Add_To_List = True Then
					swAssy_Docs.Add(New Assy_Docs(swComp_Assy, "Null", swComp.GetSelectByIDString(), RecAssyCusProp(0), RecAssyCusProp(1), 
										RecAssyCusProp(2), RecAssyCusProp(3), RecAssyCusProp(4), RecAssyCusProp(5)))
				End If

			End If
		End If

		vChildComp = swComp.GetChildren
		For i = 0 To UBound(vChildComp)

			Assy_Count = 0
			Used = False
			Dim pString = String.Empty
			Dim aString = String.Empty
			Dim RecCusProp = {"N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"}
			'RecCusProp(6)="Parent"

			swChildComp = vChildComp(i)

			If swChildComp.IsSuppressed() = False Then
				Add_To_List = False
				isAssy = False
				isPart = False
				swParent = swChildComp.GetParent

				If swParent Is Nothing Then

					swPart = swChildComp.GetModelDoc2
					pString = swChildComp.Name2



					While pString.IndexOf("/") <> -1
						pString = pString.Substring(pString.IndexOf("/") + 1)
					End While

					pString = pString.Substring(0, pString.LastIndexOf("-"))

					RecCusProp(6) = swComp_Assy

				Else


					Parent_name = swChildComp.Name2

					Assy_Count = Parent_name.Count(Function(x) x = "/")
					'MsgBox(Assy_Count)
					Debug.Print(Assy_Count.ToString + " ---- " + Parent_name)

					If Assy_Count > 1 Then
						For j = 0 To Assy_Count - 2
							Parent_name = Parent_name.Remove(0, Parent_name.IndexOf("/") + 1)
							Debug.Print(Parent_name)
						Next
						Parent_name = Parent_name.Substring(0, Parent_name.LastIndexOf("/"))
					Else
						Parent_name = Parent_name.Remove(Parent_name.LastIndexOf("/"))
					End If



					Debug.Print(Parent_name)

					'While Parent_name.IndexOf("/") <> -1
					'	Parent_name = Parent_name.Remove(Parent_name.LastIndexOf("/")
					'End While


					Parent_name = Parent_name.Substring(0, Parent_name.LastIndexOf("-"))
					Debug.Print(Parent_name)
					RecCusProp(6) = Parent_name


					swPart = swParent.GetModelDoc2
					pString = swChildComp.Name2

					While pString.IndexOf("/") <> -1
						pString = pString.Substring(pString.IndexOf("/") + 1)
					End While

					pString = pString.Substring(0, pString.LastIndexOf("-"))
				End If

				If pString IsNot "" Then

					If swPart_Docs.Count > 0 Then
						For f = 0 To swPart_Docs.Count - 1
							If swPart_Docs(f).subcomp = pString Then
								Used = True
								isAssy = False
								isPart = False
							End If
						Next
					End If

					If Used = False Then

						swPart3 = swApp.ActivateDoc3(pString, 0, 1, errorval)

						If swPart3.GetType = swDocumentTypes_e.swDocPART Then
							isPart = True
							swModelDocExt_Part = swPart3.Extension
							CusProperties_Part = swModelDocExt_Part.CustomPropertyManager("")

							For propNum = 0 To UBound(CusProp)
								Dash_Name = CusProperties_Part.Get6(CusProp(propNum), True, ValOut, RecCusProp(propNum),
													wasResolved, linkToProp)
							Next

							If CusProperties_Part.Get6(CusProp(0), True, ValOut, RecCusProp(0), wasResolved, linkToProp) = 2 Then
								If RecCusProp(0) <> "" And RecCusProp(0) <> "-XX" Then

									Temp_Name = RecCusProp(0).Substring(0, 1)

									If Temp_Name = "-" Then
										Add_To_List = True
									End If
								End If

							End If

						ElseIf swPart3.GetType = swDocumentTypes_e.swDocASSEMBLY Then
							isAssy = True
							aString = pString

							swModelDocExt_Assy = swPart3.Extension
							CusProperties_Assy = swModelDocExt_Assy.CustomPropertyManager("")

							For propNum = 0 To UBound(CusProp)
								Dash_Name = CusProperties_Assy.Get6(CusProp(propNum), True, ValOut, RecCusProp(propNum), wasResolved, 
														linkToProp)
							Next

							If CusProperties_Assy.Get6(CusProp(0), True, ValOut, RecCusProp(0), wasResolved, linkToProp) = 2 Then
								If RecCusProp(0) <> "" And RecCusProp(0) <> "-XX" Then

									Temp_Name = RecCusProp(0).Substring(0, 1)

									If Temp_Name = "-" Then
										Add_To_List = True
									End If
								End If

							End If

						End If
						swApp.CloseDoc(pString)
					End If

				End If

				If isAssy = True And Add_To_List = True Then
					isAssy = False
					If swPart.GetType = swDocumentTypes_e.swDocASSEMBLY Then
						swAssy_Docs.Add(New Assy_Docs(swComp_Assy, aString, swChildComp.GetSelectByIDString(), RecCusProp(0), RecCusProp(1),
													  RecCusProp(2), RecCusProp(3), RecCusProp(4), RecCusProp(5),
														RecCusProp(6)))
					End If
				End If

				If isPart = True And Add_To_List = True Then
					isPart = False
					If swPart_Docs.Count = 0 Then
						swPart_Docs.Add(New Part_Docs(swComp_Assy, pString, swChildComp.GetSelectByIDString(), RecCusProp(0), RecCusProp(1),
													  RecCusProp(2), RecCusProp(3), RecCusProp(4), RecCusProp(5),
														RecCusProp(6)))
					Else
						swPart_Docs.Add(New Part_Docs(swComp_Assy, pString, swChildComp.GetSelectByIDString(), RecCusProp(0), RecCusProp(1),
													  RecCusProp(2), RecCusProp(3), RecCusProp(4), RecCusProp(5), 
														RecCusProp(6)))
					End If
				End If

				Status = Add_Docs2(swChildComp, nLevel + 1)
			End If
		Next i

		Return True

	End Function

End class
