Local VLL_Return, VLO_Object, VLN_Row,VLC_Invoice,VLC_PO,VLN_TAXA, VLC_LogTXT ,VLC_String, VLC_Invoice, VLC_String, VLC_STRINGB, VLC_STRINGC, VLL_Retorno, VLL_Returnb
Store "" To VLC_Log

Store 1 To VLN_Row

&& Instanciando objeto do excell
VLO_Object = Createobject("excel.application")
VLO_Object.Workbooks.Open(This.VOC_FileName)
VLL_Retorno = Vartype(VLO_Object) == "O"

&& Criando cursor de resultados
CREATE CURSOR RESULTADO_TMP (INVOICE C(20), DATA c(10), PO C(20), TITULO C(20), VENCIMENTO c(10),;
							 MOEDA C(10), BRUTODOLAR N(8,2), ABERTODOLAR N(8,2), TAXA N(8,4), BRUTOREAL N(8,2),;
							 ABERTOREAL N(8,2), VARIACAO N(8,2), TOTAL_HARDWARE N(8,2), TOTAL_SOFTWARE N(8,2),;
							 TOTAL_SERVICO N(8,2), IRRF N(8,2), MENSAGENS C(200) )

Do While .T.
	VLC_Log = ""
	VLN_Row = VLN_Row + 1

	VLC_Invoice = Nvl(Strt(VLO_Object.Range("A"+Alltrim(Str(VLN_Row))).Text,'.',''),"")
	VLC_POInvoice = Strt(VLO_Object.Range("B"+Alltrim(Str(VLN_Row))).Text,'.','')
	If Empty(VLC_Invoice)
		Exit
	Endif

	WAIT windows "Processando invoice "+ ALLTRIM(VLC_Invoice) + " ..." nowait
	VLC_Proc = "exec [Westcon].[Starsoft].[Reports_InvoicePgto]  '" + VLC_Invoice + "','" + VLC_POInvoice + "'" + ","+ + Strtran(Alltrim(Str(This.von_taxa,12,4)),[,],[.])
	VGO_Gen.fol_sqlexec(VLC_Proc,"TMP_XLS")

	If Used("TMP_XLS") And Reccount("TMP_XLS")>0
	
		SELECT TMP_XLS
		GO TOP
		
		&& Verificando se a invoice foi lançada com o numero da PO
		IF ALLTRIM(TMP_XLS.INVOICE) == ALLTRIM(TMP_XLS.PO)
			VLC_Log = 'Invoice Lançada com Número de PO.'
		ENDIF
		
		&& Verificando se existe titulo para a invoice
		IF EMPTY(ALLTRIM(NVL(TMP_XLS.TITULO,"")))
			&& verificanco se já tem log anterior para concatenar
			IF !EMPTY(VLC_Log)
				VLC_Log = VLC_Log + ' / Não existe titulo para invoice.'
			ELSE
				VLC_Log = 'Não existe titulo para invoice.'
			endif
			
		ENDIF
		
		&& verificando e preenchendo o log no cursor temporario
		IF !EMPTY(VLC_Log)
			SELECT TMP_XLS
			replace mensagens WITH VLC_Log IN TMP_XLS
		endif
		
		SELECT TMP_XLS
		GO TOP
		scatter memvar memo
		
		SELECT RESULTADO_TMP 
		append blank
		gather memvar memo
		
		VGO_Gen.fol_closetable("TMP_XLS")
	ELSE
	
		SELECT RESULTADO_TMP 
		append blank
		replace ;
			INVOICE WITH VLC_Invoice ,;
			PO WITH VLC_POInvoice ,;
			mensagens WITH "Invoice não encontrada no sistema." ;
		IN RESULTADO_TMP 
		
	Endif
Enddo

If Used("RESULTADO_TMP")
	Select RESULTADO_TMP
	Copy To Thisform.VOC_Dir+"\InvoicePgto" Type Xls
	VGO_Gen.fol_closetable("RESULTADO_TMP")
Endif
VLO_Object.Workbooks.Close
VLO_Object.Quit



If VLL_Retorno
	VGO_Gen.FON_Msg(88)		&&-		Operação concluída
Else
	VGO_Gen.FON_Msg(689)	&&-		Operação não teve exito
Endif

VLO_Object=.Null.

Return DoDefault() And This.Cancel.Click()
