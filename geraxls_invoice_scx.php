Local VLL_Return, VLO_Object, VLN_Row,VLC_Invoice,VLC_PO,VLN_TAXA, VLC_LogTXT ,VLC_String, VLC_Invoice, VLC_String, VLC_STRINGB, VLC_STRINGC, VLL_Retorno, VLL_Returnb
Store "" To VLC_Log
*-SET STEP ON
Store 1 To VLN_Row
VLL_Returnb = .F.
VLO_Object = Createobject("excel.application")
VLO_Object.Workbooks.Open(This.VOC_FileName)
VLL_Retorno = Vartype(VLO_Object) == "O"
If VLL_Retorno
	VLL_Retorno = VGO_Gen.FOL_BeginTransaction("F11")

	If VLL_Retorno
		Do While .T.
			VLC_Log = ""
			VLN_Row = VLN_Row + 1

			VLC_Invoice = Nvl(Strt(VLO_Object.Range("A"+Alltrim(Str(VLN_Row))).Text,'.',''),"")
			VLC_POInvoice = Strt(VLO_Object.Range("B"+Alltrim(Str(VLN_Row))).Text,'.','')
			If Empty(VLC_Invoice)
				Exit
			Endif



			TEXT TO VLC_STRING TEXTMERGE NOSHOW
			SELECT U10_001_C FROM U10 (nolock)
			where U10_001_C = '<<VLC_Invoice>>'
			ENDTEXT


			TEXT TO VLC_STRINGB TEXTMERGE NOSHOW
			SELECT U10_001_C, 'Invoice Lançada com Número de PO' as LOG FROM U10 (nolock)
			where U10_001_C = '<<VLC_POInvoice>>'
			ENDTEXT


			VLL_Returnb = VGO_Gen.fol_sqlexec(VLC_STRINGB,"U10_TT")
			If VLL_Returnb And Reccount("U10_TT")>0
				VLC_Log = 'Invoice Lançada com Número de PO'
			Else
				VLL_Returnb = VGO_Gen.fol_sqlexec(VLC_String,"U10_TT")
				If VLL_Returnb And !Reccount("U10_TT")>0
					VLC_Log = 'Invoice não encontrada no sistema'
				Endif
			Endif

			VLC_Proc = "exec [Westcon].[Starsoft].[Reports_InvoicePgto]  '" + VLC_Invoice + "','" + VLC_POInvoice + "'" + ","+ + Strtran(Alltrim(Str(This.von_taxa,12,4)),[,],[.])
			VGO_Gen.fol_sqlexec(VLC_Proc,"TMP_XLS", .F.)

			If Used("TMP_XLS") And Reccount("TMP_XLS")>0
				If Used("REL") And Reccount("REL")>0
					Select *, Evl(VLC_Log,Space(200)) As Log From TMP_XLS Into Cursor rel1

					Select rel
					Insert Into rel Select * From rel1

					VGO_Gen.fol_closetable("rel1")
				Else
					Select *, Evl(VLC_Log,Space(200)) As Log From TMP_XLS Into Cursor rel Readwrite
				Endif
				VGO_Gen.fol_closetable("TMP_XLS")
			Else
				If Used("rel")
					Select rel
					Append Blank
					Replace u10_001_c With VLC_Invoice,;
						log With VLC_Log In rel
				Else
					Select *, Evl(VLC_Log,Space(200)) As Log From TMP_XLS Into Cursor rel Readwrite
					Select rel
					Append Blank
					Replace u10_001_c With VLC_Invoice,;
						log With VLC_Log In rel
				Endif
			Endif
		Enddo
	Endif
Endif
If Used("rel")
	Select rel
	Copy To Thisform.VOC_Dir+"\InvoicePgto" Type Xls
	VGO_Gen.fol_closetable("rel")
Endif
VLO_Object.Workbooks.Close
VLO_Object.Quit


If VLL_Retorno
	VGO_Gen.FOL_CommitTrans("F11")
Else
	VGO_Gen.FOL_RollBack("F11")
Endif


If VLL_Retorno
	VGO_Gen.FON_Msg(88)		&&-		Operação concluída
Else
	VGO_Gen.FON_Msg(689)	&&-		Operação não teve exito
Endif

VLO_Object=.Null.

VGO_Gen.fol_closetable("U10_TT")

Return VLL_Return And DoDefault() And This.Cancel.Click()
