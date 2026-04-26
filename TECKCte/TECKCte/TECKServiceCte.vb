Imports System.Data.Common
Imports System.IO
Imports System.Net.Mail

Public Class TECKServiceCte

    Private Const EvtLogSource As String = "TECKcte"
    Private Const EvtLogName As String = "TECKsystems"

    Public dbfactory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.DbProvider)
    Dim conn As DbConnection = Me.dbfactory.CreateConnection
    Dim cmd As DbCommand = conn.CreateCommand
    Dim fsqlBuilder As New System.Text.StringBuilder
    Private arquivoWS As StreamWriter
    Dim gcte_proxy As String = My.Settings.Proxy
    Dim gcte_proxy_usuario As String = My.Settings.Proxy_usuario
    Dim gcte_proxy_senha As String = My.Settings.Proxy_senha
    Dim gcte_vDescObrigatorio As String = My.Settings.vDescObrigatorio

    Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        'NfeBuscaCerfificado()
        ' Add any initialization after the InitializeComponent() call.
        If Not System.Diagnostics.EventLog.SourceExists(EvtLogSource) Then
            System.Diagnostics.EventLog.CreateEventSource(EvtLogSource, EvtLogName)
        End If
        EventLog1.Source = EvtLogSource
        EventLog1.WriteEntry("In OnStart") 'Escreve no log do Windows

    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try

            EventLog1.WriteEntry("In OnStart") 'Escreve no log do Windows

            arquivoWS = New StreamWriter("TeckCteService.log", True) 'Escreve no log do diretório
            arquivoWS.WriteLine("Serviço iniciado em " & DateTime.Now)
            arquivoWS.Flush()
            ' Add code here to start your service. This method should set things
            ' in motion so your service can do its work.
            conn.ConnectionString = My.Settings.ConnectionString
            conn.Open()

            'Timer que que processa as Nfes
            Timer_ctes.Interval = My.Settings.Tempo_processamento_ctes * 1000
            Timer_ctes.Enabled = True
            EventLog1.WriteEntry("Setado invervalo para processamento ctes")
            'arquivoWS.WriteLine("Setado invervalo para processamento nfes " & DateTime.Now)

            arquivoWS.WriteLine("Inicializaco timer processamento ctes " & DateTime.Now)
            arquivoWS.Flush()
        Catch ex As Exception
            'não vamos tratar exceção
        End Try

    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            Timer_ctes.Enabled = False
            arquivoWS.WriteLine("Serviço encerrado em " & DateTime.Now)
            EventLog1.WriteEntry("In OnStop")
            arquivoWS.Close()
        Catch ex As Exception
            'não vamos tratar exceção
        End Try
    End Sub

    Public Sub Processa_ctes()
        Try
            Dim conn As DbConnection = Me.dbfactory.CreateConnection
            conn.ConnectionString = My.Settings.ConnectionString
            Dim cmd As DbCommand = conn.CreateCommand
            Dim dr As DbDataReader

            Dim lsiglaWS As String

            'Pega o Nome da Máquina
            Dim lComputerName As String
            lComputerName = System.Net.Dns.GetHostName

            conn.Open()
            'Processa as NFEs                                        Fluxo Normal
            'AS - Autorizado Sefaz                                        PR
            'GP - Gerado PDF                                              GP 
            'DC - Distribuido Eletronicamente                             ES
            'ES - Enviado Sefaz                                           AS
            'IC - Impresso em Configência                                 IM
            'IM - Impresso                                                DE
            'PC - Por Cancelar
            'PI - Por Inutilizar
            'PR - Processando
            cmd.CommandText = "Select nf.dados_nfe_uf,nf.nota_fiscal,nf.numero_recibo,nf.status,nf.chave_acesso,fl.certificado_digital,nf.tipo_ambiente,nf.versao_xml,nf.filial,nf.destinatario_remetente_fornecedor_cliente,nf.emitente_cnpj,nf.impressora,nf.Dados_nfe_forma_emissao,fl.apelido_filial FROM Notas_fiscais nf, filiais fl where nf.filial = fl.filial and nf.status in ('IN','AL','IM','ES','AS','RD','PC','PI','PO') and nf.status_ultimo_evento >= getdate()-15 and nf.Dados_nfe_modelo = 57"
            dr = cmd.ExecuteReader()
            Do While dr.Read()
                Dim lstatus = dr.Item("status").ToString
                Dim lnota_fiscal As Long = dr.Item("nota_fiscal").ToString
                lsiglaWS = Função_Seleciona_siglaWS(dr.Item("dados_nfe_uf").ToString, dr.Item("Dados_nfe_forma_emissao").ToString + 1, dr.Item("Versao_xml").ToString)
                Dim lemitente_cnpj As String = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("emitente_cnpj").ToString, True, True, True, True, True, True, True, True)
                Select Case lstatus
                    Case "IN", "AL" ' Nota Incluida ou Alterada, valida e gera o XML
                        Call BDInclui_evento_nf("Em processamento", lnota_fiscal, "EP")
                        Call NFeUtil_Gera_Xml(lnota_fiscal)
                    Case "ES" 'Enviado Sefaz
                        Call NfeUtil_Consulta_Cte_Sefaz(dr.Item("certificado_digital").ToString, dr.Item("chave_acesso").ToString, Trim(dr.Item("versao_xml").ToString), lsiglaWS, dr.Item("numero_recibo").ToString, lnota_fiscal, dr.Item("tipo_ambiente").ToString, dr.Item("emitente_cnpj").ToString, "ES", dr.Item("Chave_acesso").ToString, dr.Item("dados_nfe_uf").ToString)
                    Case "AS" 'Autorizado Sefaz, agora gera o pdf para depois imprimir
                        Call Cte_gera_pdf(lnota_fiscal, lemitente_cnpj, dr.Item("chave_acesso").ToString, Trim(dr.Item("apelido_filial").ToString))
                    Case "IM" 'Impresso, agora distribuir ou redistribuir
                        If Processo_Distribuir(lnota_fiscal, dr.Item("filial").ToString, dr.Item("destinatario_remetente_fornecedor_cliente").ToString, dr.Item("chave_acesso").ToString) Then
                            If lstatus = "RD" Then
                                Call BDInclui_evento_nf("Redistribuido e´mail ao fornecedor/cliente", lnota_fiscal, "DE")
                            Else
                                Call BDInclui_evento_nf("Distribuído e´mail ao fornecedor/cliente", lnota_fiscal, "DE")
                            End If
                        Else
                            Call BDInclui_evento_nf("Não distribuído", lnota_fiscal, "ND")
                        End If
                        Call BDInclui_evento_nf("Cte finalizada", lnota_fiscal, "FN")
                    Case "PC" 'Por Cancelar
                        EventLog1.WriteEntry("Cancelando ctes " & DateTime.Now)
                        Call NfeUtil_Cancela_cte(lnota_fiscal)
                    Case "PI" 'Por Inutilizar
                        Call NfeUtil_Inutiliza_nfe(lnota_fiscal)
                    Case "PO" 'Por Corrigir
                        Call NfeUtil_carta_correcao(lnota_fiscal, BDretorna_campo("notas_fiscais_cartas_correcao", "max(numero)", "nota_fiscal = " & lnota_fiscal & " and Protocolo = ''"))
                End Select
            Loop
            dr.Close()
            conn.Close()
            cmd.Dispose()
            conn.Dispose()

        Catch err As Exception
            EventLog1.WriteEntry(err.Message)
        End Try
    End Sub

    Private Function Função_Seleciona_siglaWS(ByVal lsiglaUF As String, ByVal lDados_nfe_forma_emissao As Integer, ByVal lversao_xml As String) As String
        Função_Seleciona_siglaWS = lsiglaUF
        If lDados_nfe_forma_emissao = 1 Then 'Emissão normal
            Select Case lsiglaUF
                Case "MA", "PA", "PI"
                    Função_Seleciona_siglaWS = "SVAN"
                Case "AC", "AL", "AP", "DF", "ES", "PB", "RJ", "RN", "RO", "RR", "SC", "SE", "TO"
                    Função_Seleciona_siglaWS = "SVRS"
                Case "PR"
                    Função_Seleciona_siglaWS = "PR3"
                Case "SP"
                    If lversao_xml = "3.00" Or lversao_xml = "4.00" Then
                        Função_Seleciona_siglaWS = "SP3"
                    Else
                        Função_Seleciona_siglaWS = "SP"
                    End If
                Case "BA"
                    Função_Seleciona_siglaWS = "BA3"
                Case Else
                    Função_Seleciona_siglaWS = lsiglaUF
            End Select
        End If
        If lDados_nfe_forma_emissao = 6 Then 'Emissão contigência SVC AC, AL, AP, DF, ES, MG, PB, PI, RJ, RN, RO, RR, RS, SC, SE, SP e TO
            Função_Seleciona_siglaWS = "SVC-AN"
        End If
        If lDados_nfe_forma_emissao = 7 Then 'Emissão contigência SVC AM, BA, CE, GO, MA, MS, MT, PA, PE e PR
            Função_Seleciona_siglaWS = "SVC-RS"
        End If
    End Function

    Private Function Função_Seleciona_url_qrcode(ByVal lsiglaUF As String, ByVal ltipo_ambiente As String) As String
        Função_Seleciona_url_qrcode = ""
        If ltipo_ambiente = "1" Then ' Produção
            Select Case lsiglaUF
                Case "AC", "AL", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "PA", "PB", "PI", "RJ", "RN", "RO", "RS", "SC", "TO"
                    Função_Seleciona_url_qrcode = "https://dfe-portal.svrs.rs.gov.br/cte/qrCode"
                Case "AP", "RR", "PE", "SP"
                    Função_Seleciona_url_qrcode = "https://nfe.fazenda.sp.gov.br/CTeConsulta/qrCode"
                Case "MG"
                    Função_Seleciona_url_qrcode = "https://cte.fazenda.mg.gov.br/portalcte/sistema/qrcode.xhtml"
                Case "MS"
                    Função_Seleciona_url_qrcode = "http://www.dfe.ms.gov.br/cte/qrcode"
                Case "MT"
                    Função_Seleciona_url_qrcode = "https://www.sefaz.mt.gov.br/cte/qrcode"
                Case "PR"
                    Função_Seleciona_url_qrcode = "http://www.fazenda.pr.gov.br/cte/qrcode"
            End Select
        Else
            Select Case lsiglaUF
                Case "AC", "AL", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "PA", "PB", "PI", "RJ", "RN", "RO", "RS", "SC", "TO"
                    Função_Seleciona_url_qrcode = "https://dfe-portal.svrs.rs.gov.br/cte/qrCode"
                Case "AP", "RR", "PE", "SP"
                    Função_Seleciona_url_qrcode = "https://homologacao.nfe.fazenda.sp.gov.br/CTeConsulta/qrCode"
                Case "MG"
                    Função_Seleciona_url_qrcode = "https://hcte.fazenda.mg.gov.br/portalcte/sistema/qrcode.xhtml"
                Case "MS"
                    Função_Seleciona_url_qrcode = "http://www.dfe.ms.gov.br/cte/qrcode"
                Case "MT"
                    Função_Seleciona_url_qrcode = "https://homologacao.sefaz.mt.gov.br/cte/qrcode"
                Case "PR"
                    Função_Seleciona_url_qrcode = "http://www.fazenda.pr.gov.br/cte/qrcode"
            End Select
        End If
    End Function

    Private Sub NFeUtil_Gera_Xml(ByVal lnota_fiscal As Long)
        Dim conn As DbConnection = Me.dbfactory.CreateConnection
        conn.ConnectionString = My.Settings.ConnectionString
        Dim cmd As DbCommand = conn.CreateCommand
        Dim dr As DbDataReader
        Dim conn1 As DbConnection = Me.dbfactory.CreateConnection
        conn1.ConnectionString = My.Settings.ConnectionString
        Dim cmd1 As DbCommand = conn1.CreateCommand
        'Dim dr1 As DbDataReader

        conn.Open()
        conn1.Open()
        'Seleciona a nota fiscal
        cmd.CommandText = "Select nf.*,fl.certificado_digital,fl.Cte_chave_flexdocs,nfi.dados_cfop from notas_fiscais nf (nolock),filiais fl (nolock),Notas_fiscais_itens nfi (nolock) where nf.nota_fiscal = nfi.nota_fiscal and nf.filial = fl.filial and nf.nota_fiscal = " & lnota_fiscal
        dr = cmd.ExecuteReader()
        If dr.Read() Then
            Dim infCTeSupl_nomeCertificado As String = dr.Item("Certificado_digital").ToString
            'Dim infCarga_vCarga_Opc As Double = dr.Item("Totais_icms_total").ToString ' Valor total da carga (15 posições, sendo 13 inteiras e 2 decimais)
            Dim infCarga_vCarga_Opc As Double = CDbl(0 & dr.Item("Dados_nfe_valor_total_nota").ToString) ' Valor total da carga (15 posições, sendo 13 inteiras e 2 decimais)

            Dim objCTeUtil As Object

            objCTeUtil = CreateObject("CTe_Util.Util")
            '
            '======  Dados do  Dim Emitente do CT-e==========
            '
            Dim emi As String
            Dim emi_CNPJ As String = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Emitente_cnpj").ToString, True, True, True, True, True, True, True, True)
            Dim emi_IE As String = dr.Item("Emitente_inscricao_estadual").ToString
            Dim emi_IEST_Opc As String = dr.Item("Emitente_inscricao_estadual_substituicao_tributaria").ToString
            Dim emi_xNome As String
            If dr.Item("Tipo_ambiente").ToString = "2" Then
                emi_xNome = "CTE EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
            Else
                emi_xNome = dr.Item("Emitente_razao_social").ToString
            End If
            Dim emi_xFant_Opc As String = ""
            If dr.Item("Tipo_ambiente").ToString <> "2" Then
                emi_xFant_Opc = dr.Item("Emitente_razao_social").ToString
            End If
            Dim emi_xLgr As String = dr.Item("Emitente_endereco").ToString
            Dim emi_nro As String = dr.Item("Emitente_numero").ToString
            Dim emi_xCpl_Opc As String = dr.Item("Emitente_complemento").ToString
            Dim emi_xBairro As String = dr.Item("Emitente_bairro").ToString
            Dim emi_cMun As String = dr.Item("Emitente_municipio_codigo_ibge").ToString
            Dim emi_xMun As String = dr.Item("Emitente_municipio").ToString
            Dim emi_CEP_Opc As String = "" & dr.Item("Emitente_cep").ToString.Replace("-", "")
            Dim emi_UF As String = dr.Item("Emitente_uf").ToString
            Dim emi_fone_Opc As String = ""
            Dim emi_CRT_Opc As String = dr.Item("Emitente_regime_tributario").ToString
            '
            'emi = objCTeUtil.emitente300(emi_CNPJ, emi_IE, emi_IEST_Opc, emi_xNome, emi_xFant_Opc, emi_xLgr, emi_nro, emi_xCpl_Opc, emi_xBairro, emi_cMun, emi_xMun, emi_CEP_Opc, emi_UF, emi_fone_Opc)
            emi = objCTeUtil.emitenteCRT(emi_CNPJ, emi_IE, emi_IEST_Opc, emi_xNome, emi_xFant_Opc, emi_xLgr, emi_nro, emi_xCpl_Opc, emi_xBairro, emi_cMun, emi_xMun, emi_CEP_Opc, emi_UF, emi_fone_Opc, emi_CRT_Opc)
            '
            '======  Dados do Dim identificação do Tomador de Serviço========== Quem paga o frete
            '
            Dim toma As String
            Dim toma_toma As Long
            If dr.Item("Destinatario_remetente_fornecedor_cliente").ToString = dr.Item("Tomador_fornecedor_cliente").ToString Then
                toma_toma = 0 'Remetente
            ElseIf dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_chave").ToString = dr.Item("Tomador_fornecedor_cliente").ToString Then
                toma_toma = 1 'Expedidor
            ElseIf dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_chave").ToString = dr.Item("Tomador_fornecedor_cliente").ToString Then
                toma_toma = 2 'Recebedor
            ElseIf dr.Item("Destinatario_fornecedor_cliente").ToString = dr.Item("Tomador_fornecedor_cliente").ToString Then
                toma_toma = 3 'Destinatário
            Else
                toma_toma = 4 'Outros
            End If
            Dim toma_CPF As String = ""
            Dim toma_CNPJ As String = ""
            If Len(dr.Item("Tomador_cpf_cnpj").ToString) = 11 Then
                toma_CPF = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Tomador_cpf_cnpj").ToString, True, True, True, True, True, True, True, True) ' CPF do tomador sem máscara de formatação
            Else
                toma_CNPJ = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Tomador_cpf_cnpj").ToString, True, True, True, True, True, True, True, True) ' CNPJ do tomador sem máscara de formatação
            End If
            Dim toma_IE_Opc As String = dr.Item("Tomador_inscricao_estadual").ToString ' Inscrição Estadual do tomador sem máscara
            Dim toma_xNome As String = dr.Item("Tomador_razao_social_nome").ToString ' Razão social do tomador, evitar caracteres acentuados e &
            Dim toma_xFant_Opc As String = "" ' Nome fantasia
            Dim toma_fone_Opc As String = "" ' número do telefone sem máscara
            Dim toma_xLgr As String = dr.Item("Tomador_endereco").ToString ' logradouro
            Dim toma_nro As String = dr.Item("Tomador_numero").ToString ' número, informar S/N quano inexistente para erro de Schema XML
            Dim toma_xCpl_Opc As String = "" ' complemento do endereço, o conteúdo pode ser omitido
            Dim toma_xBairro As String = dr.Item("Tomador_bairro").ToString ' bairro
            Dim toma_cMun As String = dr.Item("Tomador_municipio_codigo_ibge").ToString ' código do município, deve ser compatível com a UF
            Dim toma_xMun As String = dr.Item("Tomador_municipio").ToString ' nome do município
            'Dim toma_CEP_Opc As String = "" ' CEP - sem máscara
            Dim toma_CEP_Opc As String = "" & dr.Item("Tomador_cep").ToString.Replace("-", "")
            Dim toma_UF As String = dr.Item("Tomador_uf").ToString ' sigla da UF
            Dim toma_cPais_Opc As String = "" ' código do pais - deve fixo em 1058 - Brasil
            Dim toma_xPais_Opc As String = "" ' nome do pais (Brasil ou BRASIL)
            Dim toma_email_Opc As String = "" ' email do tomador
            '
            toma = objCTeUtil.tomador300(toma_toma, toma_CNPJ, toma_CPF, toma_IE_Opc, toma_xNome, toma_xFant_Opc, toma_fone_Opc, toma_xLgr, toma_nro, toma_xCpl_Opc, toma_xBairro, toma_cMun, toma_xMun, toma_CEP_Opc, toma_UF, toma_cPais_Opc, toma_xPais_Opc, toma_email_Opc)
            '
            '======Identificação do documento=======
            '
            Dim identificador_cUF As Long = dr.Item("Dados_nfe_uf_ibge").ToString '35                                          ' Código da UF do emitente do CT-e
            Dim identificador_dhEmi As String
            If BDretorna_campo("Parametros", "valor_texto", "parametro = 56") = "S" Then 'Horário de Verão
                identificador_dhEmi = Format(dr.Item("Dados_nfe_data_emissao"), "yyyy-MM-ddTHH:mm:ss-02:00") ' data de emissão
            Else
                identificador_dhEmi = Format(dr.Item("Dados_nfe_data_emissao"), "yyyy-MM-ddTHH:mm:ss-03:00") ' data de emissão
            End If
            Dim identificador_mod As Long = dr.Item("Dados_nfe_modelo").ToString                                               ' Modelo do documento fiscal
            Dim identificador_serie As Long = dr.Item("Dados_nfe_serie").ToString '0                                           ' Série do CT-e
            Dim identificador_nCT As Long = dr.Item("Dados_nfe_numero").ToString '1                                            ' Número do CT-e
            Dim identificador_tpEmis As Long = dr.Item("Dados_nfe_forma_emissao").ToString + 1 '1                                 ' Forma de emissão do CT-e ( Prencher com: 1 - Normal; 5 - Contigência FSDA; 7 -  Autorização pela SVC-RS; 8 - Autorização pela SVC-SP)
            '
            '     gera a chave de acesso do CT-e
            '
            Dim cUF, ano, mes, modelo, serie, tpEmis, numero, codigoSeguranca, msgResultado, ChaveCte As String
            Dim cCT As String
            Dim cDV As String

            cUF = Trim(Str(identificador_cUF))
            ano = Mid(identificador_dhEmi, 3, 2)
            mes = Mid(identificador_dhEmi, 6, 2)
            modelo = Trim(Str(identificador_mod))
            serie = Trim(Str(identificador_serie))
            numero = Trim(Str(identificador_nCT))
            tpEmis = Trim(Str(identificador_tpEmis))
            '
            msgResultado = ""
            codigoSeguranca = "segredo"
            cCT = ""
            cDV = ""
            ChaveCte = ""

            If objCTeUtil.CriaChaveCTeNovo(cUF, ano, mes, emi_CNPJ, modelo, serie, numero, tpEmis, codigoSeguranca, msgResultado, cCT, cDV, ChaveCte) <> 0 Then
                Call BDInclui_evento_nf("CriaChaveCTeNovo: " & msgResultado, lnota_fiscal, "EE")
                Exit Sub
            End If
            cmd1.CommandText = "update notas_fiscais set Chave_acesso = '" & ChaveCte & "' where nota_fiscal = " & lnota_fiscal
            cmd1.ExecuteNonQuery()

            Dim identificador As String
            Dim identificador_cCT As Long = CLng(0 & cCT) '75                                                ' Código numérico que compões a Chave de Acesso
            Dim identificador_CFOP As String = dr.Item("Dados_cfop").ToString '"1234"                                              ' Código Fiscal de Operações e Prestações
            Dim identificador_natOp As String = dr.Item("Dados_nfe_natureza_operacao").ToString '"VENDA"                         ' Natureza da Operação
            Dim identificador_indGlobalizado As String = dr.Item("Cte_globalizado").ToString '"0"                                   ' Informar valor 1 quando for globalizado e não informar a tag nas demais situações.
            Dim identificador_tpImp As Long = "1"                                                                                 ' Formato de impressão do DACTE ( Preencher com: 1 - Retrato; 2 - Paisagem)
            Dim identificador_cDV As Long = CLng(0 & cDV) '2                                                                                ' Dígito Verificador da Chave de Acesso do CT-e
            Dim identificador_tpAmb As Long = dr.Item("Tipo_ambiente").ToString '2                                                ' Tipo de Ambiente ( Preencher com: 1 - Produção; 2 - Homologação)
            Dim identificador_tpCTe As Long = dr.Item("Dados_nfe_finalidade_emissao").ToString '0                                         ' Tipo do CT-e
            Dim identificador_procEmi As Long = 0 '0                                       ' Identificador do processo de emissão do CT-e
            Dim identificador_verProc As String = "1.2a"                                                                            ' Versão do processo de emissão
            Dim identificador_cMunEnv As String = dr.Item("Dados_nfe_municipio_codigo_ibge").ToString '"1234567"                    ' Código do Município de envio do CT-e (de onde o documento foi transmitido)
            Dim identificador_xMunEnv As String = dr.Item("Dados_nfe_municipio").ToString '"São Paulo"                              ' Nome do Município de envio do CT-e (de onde o documento foi transmitido)
            Dim identificador_UFEnv As String = dr.Item("Dados_nfe_uf").ToString '"SP"                                              ' Sigla da UF de envio do CT-e (de onde o documento foi transmitido)
            Dim identificador_modal As String = "01" '                                                 ' Modal (Preencher com: 01 - Rodoviário; 02 -  Aéreo; 03 - Aquaviário; 04 - Ferroviário; 05 - Dutoviário)
            Dim identificador_tpServ As Long = dr.Item("Cte_tipo_servico").ToString '0                                              ' Tipo de Serviço (Preencher com: 0- Normal; 1 - Subcontratação; 2 - Redespacho; 3 - Redespacho Intermediário)
            Dim identificador_cMunIni As String = dr.Item("Cte_codigo_municipio_inicio_prestacao").ToString '"1234567"              ' Código do Município de início da prestação
            Dim identificador_xMunIni As String = dr.Item("Cte_nome_municipio_inicio_prestacao").ToString '"São Paulo"              ' Nome do Município do início da prestação
            Dim identificador_UFIni As String = dr.Item("Cte_uf_municipio_inicio_prestacao").ToString '"SP"                         ' UF do início da prestação
            Dim identificador_cMunFim As String = dr.Item("Cte_codigo_municipio_fim_prestacao").ToString '"1234567"                 ' Código do Município de término da prestação
            Dim identificador_xMunFim As String = dr.Item("Cte_nome_municipio_fim_prestacao").ToString '"São Paulo"                 ' Nome do Município do término da prestação
            Dim identificador_UFFim As String = dr.Item("Cte_uf_municipio_fim_prestacao").ToString '"SP"                            ' UF do término da prestação
            Dim identificador_retira As Long = 0 '0                  ' Indicador se o Recebedor retira no Aeroporto, Filial, Porto ou Estação de Destino?
            Dim identificador_xDetRetira As String = ""                 '"Detalhes..."  ' Detalhes do retira
            Dim identificador_indIEToma As Long = dr.Item("Tomador_tipo_contribuinte").ToString '1                                       ' Indicador do papel do tomador naprestação do serviço: 1 – Contribuinte ICMS; 2 – Contribuinte isento de inscrição; 9 – Não Contribuinte.
            Dim identificador_tomador As String = toma '"..."                                   ' Tomador de Serviço, informar com o XML gerado em http://www.flexdocs.com.br/guiacte/gerarCTe.toma.html - tomador
            Dim identificador_dhCont_Opc As String = ""
            Dim identificador_xJust_Opc As String = ""
            If identificador_tpEmis > 1 Then
                If BDretorna_campo("Parametros", "valor_texto", "parametro = 56") = "S" Then 'Horário de Verão
                    identificador_dhCont_Opc = Format(dr.Item("Contigencia_data_hora"), "yyyy-MM-ddTHH:mm:ss-02:00") ' data de emissão
                Else
                    identificador_dhCont_Opc = Format(dr.Item("Contigencia_data_hora"), "yyyy-MM-ddTHH:mm:ss-03:00") ' data de emissão
                End If
                identificador_xJust_Opc = dr.Item("Contigencia_justificativa").ToString '"Web Service indisponível"       ' Justificativa da entrada em contingência
            End If
            Dim identificador_gComprasGov_Opc As String = ""
            '
            'identificador = objCTeUtil.identificador300(identificador_cUF, identificador_cCT, identificador_CFOP, identificador_natOp, identificador_mod, identificador_serie, identificador_nCT, identificador_dhEmi, identificador_tpImp, identificador_tpEmis, identificador_cDV, identificador_tpAmb, identificador_tpCTe, identificador_procEmi, identificador_verProc, identificador_indGlobalizado, identificador_cMunEnv, identificador_xMunEnv, identificador_UFEnv, identificador_modal, identificador_tpServ, identificador_cMunIni, identificador_xMunIni, identificador_UFIni, identificador_cMunFim, identificador_xMunFim, identificador_UFFim, identificador_retira, identificador_xDetRetira, identificador_indIEToma, identificador_tomador, identificador_dhCont_Opc, identificador_xJust_Opc)
            identificador = objCTeUtil.identificadorRT(identificador_cUF, identificador_cCT, identificador_CFOP, identificador_natOp, identificador_mod, identificador_serie, identificador_nCT, identificador_dhEmi, identificador_tpImp, identificador_tpEmis, identificador_cDV, identificador_tpAmb, identificador_tpCTe, identificador_procEmi, identificador_verProc, identificador_indGlobalizado, identificador_cMunEnv, identificador_xMunEnv, identificador_UFEnv, identificador_modal, identificador_tpServ, identificador_cMunIni, identificador_xMunIni, identificador_UFIni, identificador_cMunFim, identificador_xMunFim, identificador_UFFim, identificador_retira, identificador_xDetRetira, identificador_indIEToma, identificador_tomador, identificador_dhCont_Opc, identificador_xJust_Opc, identificador_gComprasGov_Opc)
            Dim licenca As String = dr.Item("Cte_chave_flexdocs").ToString
            '
            '======  Declaração dos parâmetros dos dados Complementares do CT-e==========
            '
            '
            '======  Dados do Dim Observações gerais do Contribuinte==========
            '
            Dim obsCont As String = ""
            'Dim obsCont_xCampo As String = "Observacoes"                 ' Identificação do campo
            Dim obsCont_xTexto As String = dr.Item("Informacoes_complementares_contribuinte").ToString ' Conteúdo do campo

            'obsCont = objCTeUtil.obsCont(obsCont_xCampo, obsCont_xTexto)

            '
            '======  Dados do Dim Observações gerais do Fisco==========
            '
            Dim obsFisco As String
            Dim obsFisco_xCampo As String = "10"                             ' Identificação do Campo
            Dim obsFisco_xTexto As String = ""              ' Conteúdo do campo

            obsFisco = objCTeUtil.obsFisco(obsFisco_xCampo, obsFisco_xTexto)

            Dim compl As String = ""
            Dim compl_xCaraAd_Opc As String = ""                 ' Característica Adicional do transporte (Ex: REENTREGA; DEVOLUÇÃO; REFATURAMENTO; etc)
            Dim compl_xCaraSer_Opc As String = ""                  ' Característica Adicional do serviço (Ex: ENTREGA EXPRESSA; LOGÍSTICA REVERSA; CONVENCIONAL, EMERGENCIAL; etc)
            Dim compl_xEmi_Opc As String = ""           ' Funcionário Emissor do CT-E
            Dim compl_fluxo_Opc As String = ""                         ' Previsão do Fluxo de Carga, informar com o XML gerado em fluxo
            Dim compl_entrega_Opc As String = ""                       ' Informações ref. a previsão de entrega, informar com o XML gerado em Entrega
            Dim compl_origCalc_Opc As String = ""                   ' Município de origem para efeito de cálculo do frete
            Dim compl_destCalc_Opc As String = ""                ' Município de destino para efeito de cálculo do frete
            Dim compl_xObs_Opc As String = obsCont_xTexto                             ' Observações Gerais
            Dim compl_obsCont_Opc As String = obsCont        ' Campo de Uso Livre do contribuinte - informar com o XML gerado em obsCont
            Dim compl_ObsFisco_Opc As String = obsFisco                      ' Campo de Uso Livre do contribuinte - informar com o XML gerado em obsFisco

            compl = objCTeUtil.compl(compl_xCaraAd_Opc, compl_xCaraSer_Opc, compl_xEmi_Opc, compl_fluxo_Opc, compl_entrega_Opc, compl_origCalc_Opc, compl_destCalc_Opc, compl_xObs_Opc, compl_obsCont_Opc, compl_ObsFisco_Opc)            'Dim infAdic_infAdiFisco As String = ""
            '
            '======  Dados do Dim Grupo de Previsão de Fluxo de Carga==========
            '
            Dim fluxo As String = ""
            Dim fluxo_xOrig_Opc As String = ""              ' Sigla ou código interno da Filial/Porto/Estação/Aeroporto de Origem
            Dim fluxo_pass_Opcc As String = ""                    ' Sigla ou código interno da Filial/Porto/Estação/Aeroporto de Passagem
            Dim fluxo_xDest_Opc As String = ""              ' Sigla ou código interno da Filial/Porto/Estação/Aeroporto de Destino
            Dim fluxo_xRota_Opc As String = ""                  ' Código da Rota de entrega
            '
            'fluxo = objCTeUtil.fluxo300(fluxo_xOrig_Opc, fluxo_pass_Opcc, fluxo_xDest_Opc, fluxo_xRota_Opc)
            '
            '===================grupo de Sigla ou Código interno da Filial/Porto/Estação/Aeroporto de Passagem=======================
            '
            Dim pass As String = ""
            Dim pass_xPass_Opc As String = ""              ' Sigla ou código interno da Filial/Porto/Estação/Aeroporto de Passagem
            '
            'pass = objCTeUtil.pass(pass_xPass_Opc)
            '
            '======  Dados do Dim Informações referente a previsão de Entrega==========
            '
            Dim Entrega As String
            Dim Entrega_tpPer As Long = 0 '   0-Sem data definida;'   1-Na data;'   2-Até a data;'   3-A partir da data;'   4-No período
            Dim Entrega_dIni As Date = identificador_dhEmi
            Dim Entrega_dFim As Date = identificador_dhEmi
            Dim Entrega_tpHor As Long = 0 '   0-sem hora definida;'   1-No horário; '   2-Até o horário;'   3-A partir do horário;'   4-No intervalo de tempo.
            Dim Entrega_hIni As Date = #8:00:00 AM#
            Dim Entrega_hFim As Date = #8:00:00 AM#
            '
            Entrega = objCTeUtil.Entrega(Entrega_tpPer, Entrega_dIni, Entrega_dFim, Entrega_tpHor, Entrega_hIni, Entrega_hFim)
            '
            '======  Dados do Dim remetente==========
            '
            Dim reme As String
            Dim reme_CNPJ As String = ""
            Dim reme_CPF As String = ""
            If Len(dr.Item("Destinatario_remetente_cpf_cnpj").ToString) = 18 Then
                reme_CNPJ = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_remetente_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
            Else
                reme_CPF = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_remetente_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
            End If
            Dim reme_IE_Opc As String = dr.Item("Destinatario_remetente_inscricao_estadual").ToString
            Dim reme_xNome As String = ""
            If dr.Item("Tipo_ambiente").ToString = "2" Then
                reme_xNome = emi_xNome
            Else
                reme_xNome = dr.Item("Destinatario_remetente_razao_social_nome").ToString
            End If
            Dim reme_xFant_Opc As String = ""
            Dim reme_fone_Opc As String = ""
            Dim reme_xLgr As String = dr.Item("Destinatario_remetente_endereco").ToString
            Dim reme_nro As String = dr.Item("Destinatario_remetente_numero").ToString
            Dim reme_xCpl_Opc As String = dr.Item("Destinatario_remetente_complemento").ToString
            Dim reme_xBairro As String = dr.Item("Destinatario_remetente_bairro").ToString
            Dim reme_cMun As String = dr.Item("Destinatario_remetente_municipio_codigo_ibge").ToString
            Dim reme_xMun As String = dr.Item("Destinatario_remetente_municipio").ToString
            Dim reme_CEP_Opc As String = "" & dr.Item("Destinatario_remetente_cep").ToString.Replace("-", "")
            Dim reme_UF As String = dr.Item("Destinatario_remetente_uf").ToString
            Dim reme_cPais_Opc As String = dr.Item("Destinatario_remetente_pais_codigo_ibge").ToString
            Dim reme_xPais_Opc As String = dr.Item("Destinatario_remetente_pais").ToString
            Dim reme_email_Opc As String = ""
            '
            reme = objCTeUtil.remetente300(reme_CNPJ, reme_CPF, reme_IE_Opc, reme_xNome, reme_xFant_Opc, reme_fone_Opc, reme_xLgr, reme_nro, reme_xCpl_Opc, reme_xBairro, reme_cMun, reme_xMun, reme_CEP_Opc, reme_UF, reme_cPais_Opc, reme_xPais_Opc, reme_email_Opc)            'Dim infAdic_infCPL As String = ""
            '
            '======  Dados do Dim expedidor==========
            '
            Dim exped As String = ""
            If dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_chave").ToString <> "0" Then
                Dim exped_CNPJ As String = ""
                Dim exped_CPF As String = ""
                If Len(dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_cpf_cnpj").ToString) = 18 Then
                    exped_CNPJ = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
                Else
                    exped_CPF = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
                End If
                Dim exped_IE_Opc As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_ie").ToString
                Dim exped_xNome As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_razao_social_nome").ToString
                Dim exped_fone_Opc As String = ""
                Dim exped_xLgr As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_endereco").ToString
                Dim exped_nro As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_numero").ToString
                Dim exped_xCpl_Opc As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_complemento").ToString
                Dim exped_xBairro As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_bairro").ToString
                Dim exped_cMun As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_municipio_codigo_ibge").ToString
                Dim exped_xMun As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_municipio").ToString
                Dim exped_CEP_Opc As String = ""
                Dim exped_UF As String = dr.Item("Destinatario_remetente_local_retirada_diferente_emitente_uf").ToString
                Dim exped_cPais_Opc As String = ""
                Dim exped_xPais_Opc As String = ""
                Dim exped_email_Opc As String = ""

                exped = objCTeUtil.expedidor300(exped_CNPJ, exped_CPF, exped_IE_Opc, exped_xNome, exped_fone_Opc, exped_xLgr, exped_nro, exped_xCpl_Opc, exped_xBairro, exped_cMun, exped_xMun, exped_CEP_Opc, exped_UF, exped_cPais_Opc, exped_xPais_Opc, exped_email_Opc)
            End If
            '
            '======  Dados do Dim recebedor==========
            '
            Dim receb As String = ""
            If dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_chave").ToString <> "0" Then
                Dim receb_CNPJ As String = ""
                Dim receb_CPF As String = ""
                If Len(dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_cpf_cnpj").ToString) = 18 Then
                    receb_CNPJ = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
                Else
                    receb_CPF = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
                End If
                Dim receb_IE_Opc As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_ie").ToString
                Dim receb_xNome As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_razao_social_nome").ToString
                Dim receb_fone_Opc As String = ""
                Dim receb_xLgr As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_endereco").ToString
                Dim receb_nro As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_numero").ToString
                Dim receb_xCpl_Opc As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_complemento").ToString
                Dim receb_xBairro As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_bairro").ToString
                Dim receb_cMun As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_municipio_codigo_ibge").ToString
                Dim receb_xMun As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_municipio").ToString
                Dim receb_CEP_Opc As String = ""
                Dim receb_UF As String = dr.Item("Destinatario_remetente_local_entrega_diferente_destinatario_uf").ToString
                Dim receb_cPais_Opc As String = ""
                Dim receb_xPais_Opc As String = ""
                Dim receb_email_Opc As String = ""
                '
                receb = objCTeUtil.recebedor300(receb_CNPJ, receb_CPF, receb_IE_Opc, receb_xNome, receb_fone_Opc, receb_xLgr, receb_nro, receb_xCpl_Opc, receb_xBairro, receb_cMun, receb_xMun, receb_CEP_Opc, receb_UF, receb_cPais_Opc, receb_xPais_Opc, receb_email_Opc)
            End If
            '
            '======  Dados do Dim Destinatário do CT-e==========
            '
            Dim dest As String
            Dim dest_CNPJ As String = ""
            Dim dest_CPF As String = ""
            If Len(dr.Item("Destinatario_cpf_cnpj").ToString) = 18 Then
                dest_CNPJ = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
            Else
                dest_CPF = Retira_esp_pont_hifem_colc_bar_ace_vir_igu(dr.Item("Destinatario_cpf_cnpj").ToString, True, True, True, True, True, True, True, True)
            End If
            Dim dest_IE_Opc As String = dr.Item("Destinatario_inscricao_estadual").ToString
            Dim dest_xNome As String = ""
            If dr.Item("Tipo_ambiente").ToString = "2" Then
                dest_xNome = emi_xNome
            Else
                dest_xNome = dr.Item("Destinatario_razao_social_nome").ToString
            End If
            Dim dest_fone_Opc As String = ""
            Dim dest_ISUF_Opc As String = dr.Item("Destinatario_inscricao_suframa").ToString
            Dim dest_xLgr As String = dr.Item("Destinatario_endereco").ToString
            Dim dest_nro As String = dr.Item("Destinatario_numero").ToString
            Dim dest_xCpl_Opc As String = dr.Item("Destinatario_complemento").ToString
            Dim dest_xBairro As String = dr.Item("Destinatario_bairro").ToString
            Dim dest_cMun As String = dr.Item("Destinatario_municipio_codigo_ibge").ToString
            Dim dest_xMun As String = dr.Item("Destinatario_municipio").ToString
            Dim dest_CEP_Opc As String = "" & dr.Item("Destinatario_cep").ToString.Replace("-", "")
            Dim dest_UF As String = dr.Item("Destinatario_uf").ToString
            Dim dest_cPais_Opc As String = ""
            Dim dest_xPais_Opc As String = ""
            Dim dest_email_Opc As String = ""
            '
            dest = objCTeUtil.destinatario300(dest_CNPJ, dest_CPF, dest_IE_Opc, dest_xNome, dest_fone_Opc, dest_ISUF_Opc, dest_xLgr, dest_nro, dest_xCpl_Opc, dest_xBairro, dest_cMun, dest_xMun, dest_CEP_Opc, dest_UF, dest_cPais_Opc, dest_xPais_Opc, dest_email_Opc)
            '
            '======  Dados do Dim Componentes do Valor da Prestação==========
            '
            Dim Comp As String = ""
            Dim Comp_xNome As String = ""
            Dim Comp_vComp As Double = 0
            If CDbl(dr.Item("PS1_valor").ToString) > 0 Then
                Comp_vComp = dr.Item("PS1_valor").ToString
                Comp_xNome = dr.Item("PS1_descricao").ToString
                Comp = objCTeUtil.compVPrest(Comp_xNome, Comp_vComp)
            End If
            If CDbl(dr.Item("PS2_valor").ToString) > 0 Then
                Comp_vComp = dr.Item("PS2_valor").ToString
                Comp_xNome = dr.Item("PS2_descricao").ToString
                Comp = Comp + objCTeUtil.compVPrest(Comp_xNome, Comp_vComp)
            End If
            If CDbl(dr.Item("PS3_valor").ToString) > 0 Then
                Comp_vComp = dr.Item("PS3_valor").ToString
                Comp_xNome = dr.Item("PS3_descricao").ToString
                Comp = Comp + objCTeUtil.compVPrest(Comp_xNome, Comp_vComp)
            End If

            Dim vPrest As String
            Dim vPrest_vTPrest As Double = dr.Item("Totais_icms_total_nota").ToString ' Valor Total da Prestação de Serviço (15 posições, sendo 13 inteiras e 2 decimais.
            Dim vPrest_vRec As Double = dr.Item("Totais_icms_total_produtos_servicos").ToString ' Valor a Receber (15 posições, sendo 13 inteiras e 2 decimais.)
            Dim vPrest_Comp_Opc As String = Comp ' Componente do Valor da Prestação, utilize a funcionalidade compvprest para gerar os componentes

            vPrest = objCTeUtil.vPrest(vPrest_vTPrest, vPrest_vRec, vPrest_Comp_Opc)
            '======  Dados do Dim ICMS00 - Prestação sujeito à tributação normal do ICMS==========
            '
            dr.Close()
            cmd.Dispose()
            cmd.CommandText = "Select * from Notas_fiscais_itens (nolock) where nota_fiscal = " & lnota_fiscal
            dr = cmd.ExecuteReader()
            If dr.Read() Then
            End If
            '
            'Reforma Tributária
            '
            Dim gIBSUF As String = ""
            Dim gIBSUF_pIBSUF As Double = dr.Item("RT_gIBSUF_pIBSUF").ToString
            Dim gIBSUF_pDif_Opc As Double = dr.Item("RT_gIBSUF_pDif_Opc").ToString
            Dim gIBSUF_vDif_Opc As Double = dr.Item("RT_gIBSUF_vDif_Opc").ToString
            Dim gIBSUF_vDevTrib_Opc As Double = dr.Item("RT_gIBSUF_vDevTrib_Opc").ToString
            Dim gIBSUF_pRedAliq_Opc As Double = dr.Item("RT_gIBSUF_pRedAliq_Opc").ToString
            Dim gIBSUF_pAliqEfet_Opc As Double = dr.Item("RT_gIBSUF_pAliqEfet_Opc").ToString
            Dim gIBSUF_vIBSUF As Double = dr.Item("RT_gIBSUF_vIBSUF").ToString
            gIBSUF = objCTeUtil.gIBSUF(gIBSUF_pIBSUF, gIBSUF_pDif_Opc, gIBSUF_vDif_Opc, gIBSUF_vDevTrib_Opc, gIBSUF_pRedAliq_Opc, gIBSUF_pAliqEfet_Opc, gIBSUF_vIBSUF)

            Dim gIBSMun As String
            Dim gIBSMun_pIBSMun As Double = dr.Item("RT_gIBSMun_pIBSMun").ToString
            Dim gIBSMun_pDif_Opc As Double = dr.Item("RT_gIBSMun_pDif_Opc").ToString
            Dim gIBSMun_vDif_Opc As Double = dr.Item("RT_gIBSMun_vDif_Opc").ToString
            Dim gIBSMun_vDevTrib_Opc As Double = dr.Item("RT_gIBSMun_vDevTrib_Opc").ToString
            Dim gIBSMun_pRedAliq_Opc As Double = dr.Item("RT_gIBSMun_pRedAliq_Opc").ToString
            Dim gIBSMun_pAliqEfet_Opc As Double = dr.Item("RT_gIBSMun_pAliqEfet_Opc").ToString
            Dim gIBSMun_vIBSMun As Double = dr.Item("RT_gIBSMun_vIBSMun").ToString
            gIBSMun = objCTeUtil.gIBSMun(gIBSMun_pIBSMun, gIBSMun_pDif_Opc, gIBSMun_vDif_Opc, gIBSMun_vDevTrib_Opc, gIBSMun_pRedAliq_Opc, gIBSMun_pAliqEfet_Opc, gIBSMun_vIBSMun)
            Dim vIBS_UF_Mun As Double = gIBSUF_vIBSUF + gIBSMun_vIBSMun

            Dim gCBS As String
            Dim gCBS_pCBS As Double = dr.Item("RT_gCBS_pCBS").ToString
            Dim gCBS_pDif_Opc As Double = dr.Item("RT_gCBS_pDif_Opc").ToString
            Dim gCBS_vDif_Opc As Double = dr.Item("RT_gCBS_vDif_Opc").ToString
            Dim gCBS_vDevTrib_Opc As Double = dr.Item("RT_gCBS_vDevTrib_Opc").ToString
            Dim gCBS_pRedAliq_Opc As Double = dr.Item("RT_gCBS_pRedAliq_Opc").ToString
            Dim gCBS_pAliqEfet_Opc As Double = dr.Item("RT_gCBS_pAliqEfet_Opc").ToString
            Dim gCBS_vIBSMun As Double = dr.Item("RT_gCBS_vCBS").ToString
            gCBS = objCTeUtil.gCBS(gCBS_pCBS, gCBS_pDif_Opc, gCBS_vDif_Opc, gCBS_vDevTrib_Opc, gCBS_pRedAliq_Opc, gCBS_pAliqEfet_Opc, gCBS_vIBSMun)

            Dim vBC As Double = dr.Item("Dados_valor_total_bruto").ToString
            Dim vBCIBSCBS As Double
            vBCIBSCBS = vBC
            Dim gTribRegular_Opc As String = ""
            Dim gTribCompraGov_Opc As String = ""
            Dim IBSCBS As String
            IBSCBS = objCTeUtil.gIBSCBSv110(vBCIBSCBS, gIBSUF, gIBSMun, vIBS_UF_Mun, gCBS, gTribRegular_Opc, gTribCompraGov_Opc)
            '
            'Final Reforma Tributária
            '
            Dim ImpICMS00 As String = ""
            Dim vTotDFe_Opc As Double
            If dr.Item("Icms_situacao_tributaria").ToString = "00" Then
                Dim ImpICMS00_CST As String = "00" ' Classificação Tributária do Serviço (00 - tributação normal ICMS)
                Dim ImpICMS00_vBC As Double = dr.Item("Icms_base_calculo").ToString ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS00_pICMS As Double = dr.Item("Icms_aliquota").ToString ' Valor do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS00_vICMS As Double = dr.Item("Icms_valor").ToString ' Valor do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS00_vTotTrib_Opc As Double = 0 '  valor Total dos Tributos Aproximado
                Dim ImpICMS00_infAdFisco_Opc As String = "" ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
                Dim ImpICMS00_ICMSUFFim_Opc As String = "" ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
                '
                'ImpICMS00 = objCTeUtil.ImpICMS00_NT2015003(ImpICMS00_CST, ImpICMS00_vBC, ImpICMS00_pICMS, ImpICMS00_vICMS, ImpICMS00_vTotTrib_Opc, ImpICMS00_infAdFisco_Opc, ImpICMS00_ICMSUFFim_Opc)
                ImpICMS00 = objCTeUtil.ImpCST00(ImpICMS00_CST, ImpICMS00_vBC, ImpICMS00_pICMS, ImpICMS00_vICMS, ImpICMS00_vTotTrib_Opc, ImpICMS00_infAdFisco_Opc, ImpICMS00_ICMSUFFim_Opc, IBSCBS, vTotDFe_Opc)
            End If
            '
            '======  Dados do Dim ICMS20 - Prestação sujeito à tributação com redução da BC do ICMS==========
            '
            Dim ImpICMS20 As String = ""
            If dr.Item("Icms_situacao_tributaria").ToString = "20" Then
                Dim ImpICMS20_CST As String = "20" ' Classificação Tributária do Serviço (20 - tributação com BC reduzida do ICMS)
                Dim ImpICMS20_pRedBC As Double = dr.Item("Icms_percentual_reducao_base_calculo").ToString ' Percentual da redução da BC (5 posições sendo 3 inteiras e 2 decimais)
                Dim ImpICMS20_vBC As Double = dr.Item("Icms_base_calculo").ToString ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS20_pICMS As Double = dr.Item("Icms_aliquota").ToString ' Alíquota do ICMS (5 posições sendo 3 inteiras e 2 decimais)"
                Dim ImpICMS20_vICMS As Double = dr.Item("Icms_valor").ToString ' Valor do ICMS (15 posições, sendo 13 inteiras e 2 decimais)"
                Dim ImpICMS20_vTotTrib_Opc As Double = 0 ' valor Total dos Tributos Aproximado
                Dim ImpICMS20_infAdFisco_Opc As String = "" ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
                Dim ImpICMS20_ICMSUFFim_Opc As String = "" ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
                '
                ImpICMS20 = objCTeUtil.ImpCST20(ImpICMS20_CST, ImpICMS20_pRedBC, ImpICMS20_vBC, ImpICMS20_pICMS, ImpICMS20_vICMS, ImpICMS20_vTotTrib_Opc, ImpICMS20_infAdFisco_Opc, ImpICMS20_ICMSUFFim_Opc, IBSCBS, vTotDFe_Opc)
            End If
            '
            '======  Dados do Dim ICMS45 - ICMS Isento, não Tributado ou diferido==========
            '
            Dim ImpICMS45 As String = ""
            If dr.Item("Icms_situacao_tributaria").ToString = "40" Or dr.Item("Icms_situacao_tributaria").ToString = "41" Or dr.Item("Icms_situacao_tributaria").ToString = "51" Then
                Dim ImpICMS45_CST As String = dr.Item("Icms_situacao_tributaria").ToString ' = "40" ' Classificação Tributária do Serviço (40 - ICMS isenção; 41 - ICMS não tributada; 51 - ICMS diferido)
                Dim ImpICMS45_vTotTrib_Opc As Double = 0 ' valor Total dos Tributos Aproximado
                Dim ImpICMS45_infAdFisco_Opc As String = "" ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
                Dim ImpICMS45_ICMSUFFim_Opc As String = "" ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
                Dim vICMSDeson_Opc As Double = 0
                Dim cBenef_Opc As String = ""
                ImpICMS45 = objCTeUtil.ImpCST45(ImpICMS45_CST, ImpICMS45_vTotTrib_Opc, ImpICMS45_infAdFisco_Opc, ImpICMS45_ICMSUFFim_Opc, vICMSDeson_Opc, cBenef_Opc, IBSCBS, vTotDFe_Opc)
            End If
            '
            '======  Dados do Dim Tributação pelo ICMS 60 - ICMS cobrado por substituição tributária==========
            '
            Dim ImpICMS60 As String = ""
            If dr.Item("Icms_situacao_tributaria").ToString = "60" Then
                Dim ImpICMS60_CST As String = "60" ' Classificação Tributária do Serviço (60 - ICMS cobrado anteriormente por substituição tributária)
                'Dim ImpICMS60_vBCSTRet As Double = dr.Item("Icms_substituicao_tributaria_base").ToString ' Valor da BC do ICMS ST retido (15 posições, sendo 13 inteiras e 2 decimais)
                'Dim ImpICMS60_pICMSSTRet As Double = dr.Item("Icms_substituicao_tributaria_aliquota").ToString ' Valor do ICMS ST retido (15 posições, sendo 13 inteiras e 2 decimais)
                'Dim ImpICMS60_vICMSSTRet As Double = dr.Item("Icms_substituicao_tributaria_retido").ToString ' Alíquota do ICMS (5 posições sendo 3 inteiras e 2 decimais)
                Dim ImpICMS60_vBCSTRet As Double = dr.Item("Icms_base_calculo").ToString ' Valor da BC do ICMS ST retido (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS60_pICMSSTRet As Double = dr.Item("Icms_valor").ToString ' Valor do ICMS ST retido (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS60_vICMSSTRet As Double = dr.Item("Icms_aliquota").ToString ' Alíquota do ICMS (5 posições sendo 3 inteiras e 2 decimais)
                Dim ImpICMS60_vCred As Double = 0 'dr.Item("Icms_credito_aproveitado").ToString ' Valor do Crédito outorgado/Presumido (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS60_vTotTrib_Opc As Double = 0 ' valor Total dos Tributos Aproximado
                Dim ImpICMS60_infAdFisco_Opc As String = "" ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
                Dim ImpICMS60_ICMSUFFim_Opc As String = "" ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
                Dim vICMSDeson_Opc As Double = 0
                Dim cBenef_Opc As String = ""
                ImpICMS60 = objCTeUtil.ImpCST60(ImpICMS60_CST, ImpICMS60_vBCSTRet, ImpICMS60_pICMSSTRet, ImpICMS60_vICMSSTRet, ImpICMS60_vCred, ImpICMS60_vTotTrib_Opc, ImpICMS60_infAdFisco_Opc, ImpICMS60_ICMSUFFim_Opc, vICMSDeson_Opc, cBenef_Opc, IBSCBS, vTotDFe_Opc)
            End If
            '
            '======  Dados do Dim ICMS 90 - ICMS Outros==========
            '
            Dim ImpICMS90 As String = ""
            If dr.Item("Icms_situacao_tributaria").ToString = "90" Then ' And identificador_UFFim = identificador_UFIni Then
                Dim ImpICMS90_CST As String = "90" ' Classificação Tributária do Serviço (90 - ICMS outros)
                Dim ImpICMS90_pRedBC As Double = dr.Item("Icms_percentual_reducao_base_calculo").ToString ' Percentual de redução da BC (5 posições sendo 3 inteiras e 2 decimais)
                Dim ImpICMS90_vBC As Double = dr.Item("Icms_base_calculo").ToString ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS90_pICMS As Double = dr.Item("Icms_aliquota").ToString ' Alíquota do ICMS (5 posições sendo 3 inteiras e 2 decimais)
                Dim ImpICMS90_vICMS As Double = dr.Item("Icms_valor").ToString ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS90_vCred As Double = dr.Item("Icms_credito_aproveitado").ToString ' Valor do Crédito outorgado/Presumido (15 posições, sendo 13 inteiras e 2 decimais)
                Dim ImpICMS90_vTotTrib_Opc As Double = 0 ' valor Total dos Tributos Aproximado
                Dim ImpICMS90_infAdFisco_Opc As String = "" ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
                Dim ImpICMS90_ICMSUFFim_Opc As String = "" ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
                Dim vICMSDeson_Opc As Double = 0
                Dim cBenef_Opc As String = ""
                ImpICMS90 = objCTeUtil.ImpCST90(ImpICMS90_CST, ImpICMS90_pRedBC, ImpICMS90_vBC, ImpICMS90_pICMS, ImpICMS90_vICMS, ImpICMS90_vCred, ImpICMS90_vTotTrib_Opc, ImpICMS90_infAdFisco_Opc, ImpICMS90_ICMSUFFim_Opc, vICMSDeson_Opc, cBenef_Opc, IBSCBS, vTotDFe_Opc)
            End If
            '
            '======  Dados do Dim ICMS Outra UF - ICMS devido à UF de origem da prestação==========
            '
            Dim ImpICMSOutraUF As String = ""
            'If identificador_UFFim <> identificador_UFIni And dr.Item("Icms_situacao_tributaria").ToString = "90" Then
            'Dim ImpICMSOutraUF_CST As String = "90" ' Classificação Tributária do Serviço (90 - ICMS outros)
            'Dim ImpICMSOutraUF_pRedBCOutraUF As Double = 0 ' Percentual de redução da BC (5 posições sendo 3 inteiras e 2 decimais)
            'Dim ImpICMSOutraUF_vBCOutraUF As Double = 0 ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
            'Dim ImpICMSOutraUF_pICMSOutraUF As Double = 0 ' Alíquota do ICMS (5 posições sendo 3 inteiras e 2 decimais)
            'Dim ImpICMSOutraUF_vICMSOutraUF As Double = 0 ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
            'ImpICMSOutraUF_pRedBCOutraUF = dr.Item("Icms_percentual_reducao_base_calculo").ToString ' Percentual de redução da BC (5 posições sendo 3 inteiras e 2 decimais)
            'ImpICMSOutraUF_vBCOutraUF = dr.Item("Icms_base_calculo").ToString ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
            'ImpICMSOutraUF_pICMSOutraUF = dr.Item("Icms_aliquota").ToString ' Alíquota do ICMS (5 posições sendo 3 inteiras e 2 decimais)
            'ImpICMSOutraUF_vICMSOutraUF = dr.Item("Icms_valor").ToString ' Valor da BC do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
            'Dim ImpICMSOutraUF_vTotTrib_Opc As Double = 0 ' valor Total dos Tributos Aproximado
            'Dim ImpICMSOutraUF_infAdFisco_Opc As String = "" ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
            'Dim ImpICMSOutraUF_ICMSUFFim_Opc As String = "" ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
            '
            'ImpICMSOutraUF = objCTeUtil.ImpICMSOutraUF_NT2015003(ImpICMSOutraUF_CST, ImpICMSOutraUF_pRedBCOutraUF, ImpICMSOutraUF_vBCOutraUF, ImpICMSOutraUF_pICMSOutraUF, ImpICMSOutraUF_vICMSOutraUF, ImpICMSOutraUF_vTotTrib_Opc, ImpICMSOutraUF_infAdFisco_Opc, ImpICMSOutraUF_ICMSUFFim_Opc)
            'End If
            Dim imp As String = ImpICMS00 + ImpICMS20 + ImpICMS45 + ImpICMS60 + ImpICMS90 + ImpICMSOutraUF
            dr.Close()
            cmd.Dispose()
            '
            '======Dados do Dim ICMSSN - Simples Nacional==========
            '
            Dim ImpICMSSN As String = ""
            'Dim ImpICMSSN_vTotTrib_Opc As Double = dr.Item("Emitente_cnpj").ToString ' valor Total dos Tributos Aproximado
            'Dim ImpICMSSN_infAdFisco_Opc As String = dr.Item("Emitente_cnpj").ToString ' Informações adicionais de interesse do Fisco (Norma referenciada, informações complementares, etc)
            'Dim ImpICMSSN_ICMSUFFim_Opc As String = dr.Item("Emitente_cnpj").ToString ' Informações do ICMS devido para a UF de término do serviço de transporte, nas operações interestaduais para consumidor final
            '
            'ImpICMSSN = objCTeUtil.ImpICMSSN300(ImpICMSSN_vTotTrib_Opc, ImpICMSSN_infAdFisco_Opc, ImpICMSSN_ICMSUFFim_Opc)
            '
            '=======declaração de parâmetros========
            '
            Dim ICMSUFFIM As String = ""
            'Dim vBCUFFim As Double = dr.Item("Emitente_cnpj").ToString
            'Dim pICMSUFFim As Double = dr.Item("Emitente_cnpj").ToString
            'Dim pFCPUFFim As Double = dr.Item("Emitente_cnpj").ToString
            'Dim pICMSInter As Double = dr.Item("Emitente_cnpj").ToString
            'Dim pICMSInterPart As Double = dr.Item("Emitente_cnpj").ToString
            'Dim vFCPUFFim As Double = dr.Item("Emitente_cnpj").ToString
            'Dim vICMSUFFim As Double = dr.Item("Emitente_cnpj").ToString
            'Dim vICMSUFIni As Double = dr.Item("Emitente_cnpj").ToString

            'ICMSUFFIM = objCTeUtil.ICMSUFFIM(vBCUFFim, pICMSUFFim, pFCPUFFim, pICMSInter, pICMSInterPart, vFCPUFFim, vICMSUFFim, vICMSUFIni)
            '
            '======  Dados do Dim do Grupo de Informações de quantidades de Carga do CT-e==========
            '
            Dim infQ As String = ""
            cmd.CommandText = "Select * from Notas_fiscais_volumes (nolock) where nota_fiscal = " & lnota_fiscal
            dr = cmd.ExecuteReader()
            If dr.Read() Then
            End If
            Dim infQ_cUnid As String = "01" ' ' Código da Unidade de Medida (00 - M3; 01 - KG; 02 - TON; 03 - UNIDADE; 04 - LITROS; 05 - MMBTU)
            Dim infQ_tpMed As String = "PESO LIQUIDO" ' Exemplos: PESO BRUTO, PESO DELCARADO, PESO CUBADO, PESO AFORADO, PESO AFERIDO, PESO BASE DE CÁLCULO, LITRAGEM, CAIXAS, etc."
            Dim infQ_qCarga As Double = Trim(dr.Item("Quantidade").ToString) '15 posiçõies, sendo 11 inteiras e 4 decimais
            '
            infQ = objCTeUtil.infQ(infQ_cUnid, infQ_tpMed, infQ_qCarga)
            '
            '======  Dados do Dim do Grupo de Informações da Carga do CT-e==========
            '
            Dim infCarga As String = ""
            Dim infCarga_prodPred As String = Trim(dr.Item("Especie").ToString) ' Produto predominante
            Dim infCarga_xOutCat_Opc As String = Trim(dr.Item("Marca").ToString) ' Outras caracterísiticas da carga (Ex: FRIA; GRANEL; REFRIGERADA; Medidas:12X12X12
            Dim infCarga_infQ As String = infQ 'Informações de quantidade da Carga do CT-e (1 - Peso Bruto, sempre em quilogramas; 2 - Peso Cubado, sempre em quilogramas; 3 - Quantidades de volumes, sempre em unidades)
            Dim infCarga_vCarga_Averb_Opc As Double = 0 ' Valor total da carga para efeito de averbação (15 posições, sendo 13 inteiras e 2 decimais)
            dr.Close()
            cmd.Dispose()
            '
            infCarga = objCTeUtil.infCarga300(infCarga_vCarga_Opc, infCarga_prodPred, infCarga_xOutCat_Opc, infCarga_infQ, infCarga_vCarga_Averb_Opc)
            '
            '======  Dados do Dim Informações das NF das mercadorias transportadas pelo CT-e==========
            '
            Dim infNF As String = ""
            If identificador_tpCTe <> 1 Then
                cmd.CommandText = "Select * from Notas_fiscais_referenciadas_produtor (nolock) where nota_fiscal = " & lnota_fiscal
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    Dim infNF_nRoma_Opc As String = "" ' Número do Romaneio da NF
                    Dim infNF_nPed_Opc As String = "" ' Número do pedido da NF
                    Dim infNF_mod As String = dr.Item("Modelo").ToString ' Modelo da Nota Fiscal. Preencher com: 01 - NF Modelo 01/1A e Avulsa;  04 - NF de Produtor
                    Dim infNF_serie As String = dr.Item("Serie").ToString ' Série
                    Dim infNF_nDoc As String = dr.Item("Numero").ToString ' Número do Documento
                    Dim infNF_dEmi As Date = String.Format("{0:yyyy-MM-dd}", dr.Item("Mes_ano_emissao").ToString) ' Data de emissão da NF (Formato AAAA-MM-DD)
                    Dim infNF_vBC As Double = dr.Item("Icms_base_calculo").ToString ' Valor da Base de Cálculo do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                    Dim infNF_vICMS As Double = dr.Item("Icms_valor").ToString ' Valor total do ICMS (15 posições, sendo 13 inteiras e 2 decimais)
                    Dim infNF_vBCST As Double = dr.Item("Icms_base_calculo_st").ToString ' Valor da Base de Cálculo do ICMS ST (15 posições, sendo 13 inteiras e 2 decimais)
                    Dim infNF_vST As Double = dr.Item("Icms_valor_st").ToString ' Valor total do ICMS ST (15 posições, sendo 13 inteiras e 2 decimais)
                    Dim infNF_vProd As Double = dr.Item("Valor_total_produtos").ToString ' Valor total dos produtos (15 posições, sendo 13 inteiras e 2 decimais)
                    Dim infNF_vNF As Double = dr.Item("Valor_nota_fiscal").ToString ' Valor total da NF (15 posições, sendo 13 inteiras e 2 decimais)
                    Dim infNF_nCFOP As String = dr.Item("cfop").ToString ' CFOP Predominante (CFOP da NF ou, na existência de mais de um, predominância pelo critério de valor econômico)
                    Dim infNF_nPeso_Opc As Double = 0 ' Peso total em KG (15 posições, sendo 12 inteiras e 3 decimais)
                    Dim infNF_PIN_Opc As String = "" ' PIN SUFRAMA (PIN atribuído pela SUFRAMA para a operação)
                    Dim infNF_dPrevOpc As Date = DateTime.MinValue
                    Dim infNF_unidCargaTransp_Opc As String = ""
                    '
                    infNF = infNF + objCTeUtil.infNF_2G(infNF_nRoma_Opc, infNF_nPed_Opc, infNF_mod, infNF_serie, infNF_nDoc, infNF_dEmi, infNF_vBC, infNF_vICMS, infNF_vBCST, infNF_vST, infNF_vProd, infNF_vNF, infNF_nCFOP, infNF_nPeso_Opc, infNF_PIN_Opc, infNF_dPrevOpc, infNF_unidCargaTransp_Opc)
                Loop
                dr.Close()
                cmd.Dispose()
            End If
            '
            '======  Dados do Dim Informações das NF-e das mercadorias transportadas pelo CT-e==========
            '
            Dim infNFe As String = ""
            If identificador_tpCTe <> 1 Then
                cmd.CommandText = "Select * from Notas_fiscais_nfes_referenciados (nolock) where nota_fiscal = " & lnota_fiscal
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    Dim infNFe_chave As String = dr.Item("Chave_nfe").ToString ' Número da Chave de Acesso das NF-e
                    Dim infNFe_PIN_Opc As String = "" 'PIN SUFRAMA (PIN atribuído pela SUFRAMA para a operação)
                    Dim infNFe_dPrevOpc As Date = "#12:00:00 AM#"
                    Dim infNFe_unidCargaTransp_Opc As String = "" '""
                    '
                    infNFe = infNFe + objCTeUtil.infNFe_2G(infNFe_chave, infNFe_PIN_Opc, infNFe_dPrevOpc, infNFe_unidCargaTransp_Opc)
                Loop
                dr.Close()
                cmd.Dispose()
            End If
            '
            '======  Dados do Dim de Informações dos demais documentos==========
            '
            Dim infOutros As String = ""
            'Dim infOutros_tpDoc As String = "" ' Tipo de documento originário. (Preencher com: 00 - Declaração; 10 - Dutoviário; 99 - Outros)
            'Dim infOutros_descOutros_Opc As String = "" ' Descrição quando se trata de 99 - Outros
            'Dim infOutros_nDoc_Opc As String = "" ' Número do documento
            'Dim infOutros_dEmi_Opc As Date = "" ' Data de emissão (Formato AAAA-MM-DD)
            'Dim infOutros_vDocFisc_Opc As Double = 0 ' Valor do documento (15 posições, sendo 13 inteiras e 2 decimais.)
            'Dim infOutros_dPrevOpc As Date = "" '#12:00:00 AM#
            'Dim infOutros_unidCargaTransp_Opc As String = ""
            '
            'infOutros = objCTeUtil.infOutros_2G(infOutros_tpDoc, infOutros_descOutros_Opc, infOutros_nDoc_Opc, infOutros_dEmi_Opc, infOutros_vDocFisc_Opc, infOutros_dPrevOpc, infOutros_unidCargaTransp_Opc)
            '
            '======  Dados do Dim do Grupo de Informações da unidade de transporte==========
            '
            cmd.CommandText = "Select * from Notas_fiscais_reboques (nolock) where nota_fiscal = " & lnota_fiscal
            dr = cmd.ExecuteReader()
            Dim infUnidTransp As String = ""
            Dim rod_RNTRC As String = "" ' Registro Nacional de Transportadores Rodoviários de Cargas
            If dr.Read() Then
                Dim infUnidTransp_tpUnidTransp As String = dr.Item("Tipo").ToString ' Tipo da Unidade de Transporte
                Dim infUnidTransp_idUnidTransp As String = dr.Item("Placa").ToString
                If IsNumeric(dr.Item("Rtnc").ToString) Then
                    'rod_RNTRC = dr.Item("Rtnc").ToString.PadLeft(8 - Len(dr.Item("Rtnc").ToString), "0")
                    rod_RNTRC = String.Format("{0:D8}", CLng(dr.Item("Rtnc").ToString))
                Else
                    rod_RNTRC = dr.Item("Rtnc").ToString
                End If
                ' Registro Nacional de Transportadores Rodoviários de Cargas
                Dim infUnidTransp_lacUnidTransp_Opc As String = "" ' lacres da unidades de transporte se houver
                Dim infUnidTransp_infUnidCarga_Opc As String = "" ' dispositivo de carga utilizada (Unit Load Device - ULD)
                Dim infUnidTransp_qtdRat_Opc As String = "" ' quantidade rateada
                infUnidTransp = objCTeUtil.infUnidTransp(infUnidTransp_tpUnidTransp, infUnidTransp_idUnidTransp, infUnidTransp_lacUnidTransp_Opc, infUnidTransp_infUnidCarga_Opc, infUnidTransp_qtdRat_Opc)
            End If
            dr.Close()
            cmd.Dispose()
            '
            '======  Dados do Dim de número de Lacre==========
            '
            Dim lacUnidTransp As String = ""
            'Dim lacUnidTransp_nlacre As String = "" ' Número do Lacre
            '
            'lacUnidTransp = objCTeUtil.lacreUnidTransp(lacUnidTransp_nlacre)
            '
            '======  Dados do Dim do Grupo de Informações da unidade de Carga==========
            '
            Dim infUnidCarga As String = ""
            Dim infUnidCarga_tpUnidCarga As String = "4" ' Tipo da Unidade de Carga
            Dim infUnidCarga_idUnidCarga As String = infUnidTransp ' identificação da unidade de Carga
            Dim infUnidCarga_lacUnidCarga_Opc As String = "" ' lacres da unidades de Carga se houver
            Dim infUnidCarga_qtdRat_Opc As String = "" ' quantidade rateada
            '
            infUnidCarga = objCTeUtil.infUnidCarga(infUnidCarga_tpUnidCarga, infUnidCarga_idUnidCarga, infUnidCarga_lacUnidCarga_Opc, infUnidCarga_qtdRat_Opc)
            '
            '======  Dados do Dim de número de Lacre==========
            '
            Dim lacUnidCarga As String = ""
            'Dim lacUnidCarga_nlacre As String = "" ' Número do Lacre
            '
            'lacUnidCarga = objCTeUtil.lacreUnidCarga(lacUnidCarga_nlacre)
            '
            '========Dados do Dim Emissor do Documento Anterior==========
            '
            Dim EmiDocAnt As String = ""
            Dim idDoc1Ant As String = ""
            Dim idDoc2Ant As String = ""
            cmd.CommandText = "Select * from Notas_fiscais_documentos_anteriores (nolock) where nota_fiscal = " & lnota_fiscal
            dr = cmd.ExecuteReader()
            Do While dr.Read()

                Dim EmiDocAnt_CNPJ As String = dr.Item("Cnpj").ToString ' CNPJ do Emissor do documento anterior sem máscara de formatação
                Dim EmiDocAnt_CPF As String = dr.Item("Cpf").ToString ' CPF do Emissor do documento anterior, uso exclusivo do Fisco
                Dim EmiDocAnt_IE_Opc As String = dr.Item("IE").ToString ' Inscrição Estadual do Emissor do documento anterior sem máscara
                Dim EmiDocAnt_UF_Opc As String = dr.Item("UF").ToString ' sigla da UF
                Dim EmiDocAnt_xNome As String = dr.Item("Nome").ToString ' Razão social do EmiDocAntdor, evitar caracteres acentuados e &

                Dim EmiDoc1_TipoDoc As String = dr.Item("Doc1_TipoDoc").ToString ' Preencher com: 00-CTRC; 01-CTAC; 02-ACT; 03-NF Modelo 7; 04-NF Modelo 27; 05-Conhecimento Aéreo Nacional; 06-CTMC; 07-ATRE; 08-DTA(Despacho de Transito Aduaneiro); 09-Conhecimento Aereo Iternacional; 10-Conhecimento-Carta de Porte Internacional; 11-Conhecimento Avulso; 12-TIF(Transporte Internacional Ferroviário); 99-Outros
                Dim EmiDoc1_Serie As String = dr.Item("Doc1_Serie").ToString ' Serie do Documento Fiscal
                Dim EmiDoc1_SubSerie As String = dr.Item("Doc1_SubSerie").ToString ' SubSerie do Documento Fiscal
                Dim EmiDoc1_Numero As String = dr.Item("Doc1_Numero").ToString ' Número do Documento Fiscal
                Dim EmiDoc1_DataEmissao As String = dr.Item("Doc1_DataEmissao").ToString ' Data de Emissão
                Dim EmiDoc1_ChaveCte As String = dr.Item("Doc1_chavecte").ToString ' Chave Cte Eletrônico

                If Len(Trim(EmiDoc1_ChaveCte)) <> 44 And Trim(EmiDoc1_TipoDoc) <> "" And Trim(EmiDoc1_Serie) <> "" And Trim(EmiDoc1_SubSerie) <> "" And Trim(EmiDoc1_Numero) <> "" And Trim(EmiDoc1_DataEmissao) <> "" Then
                    idDoc1Ant = objCTeUtil.idDocAntPap(EmiDoc1_TipoDoc, EmiDoc1_Serie, EmiDoc1_SubSerie, EmiDoc1_Numero, EmiDoc1_DataEmissao)
                Else
                    idDoc1Ant = objCTeUtil.idDocAntEle300(EmiDoc1_ChaveCte)
                End If
                Dim EmiDoc2_TipoDoc As String = dr.Item("Doc2_TipoDoc").ToString ' Preencher com: 00-CTRC; 01-CTAC; 02-ACT; 03-NF Modelo 7; 04-NF Modelo 27; 05-Conhecimento Aéreo Nacional; 06-CTMC; 07-ATRE; 08-DTA(Despacho de Transito Aduaneiro); 09-Conhecimento Aereo Iternacional; 10-Conhecimento-Carta de Porte Internacional; 11-Conhecimento Avulso; 12-TIF(Transporte Internacional Ferroviário); 99-Outros
                Dim EmiDoc2_Serie As String = dr.Item("Doc2_Serie").ToString ' Serie do Documento Fiscal
                Dim EmiDoc2_SubSerie As String = dr.Item("Doc2_SubSerie").ToString ' SubSerie do Documento Fiscal
                Dim EmiDoc2_Numero As String = dr.Item("Doc2_Numero").ToString ' Número do Documento Fiscal
                Dim EmiDoc2_DataEmissao As String = dr.Item("Doc2_DataEmissao").ToString ' Data de Emissão
                Dim EmiDoc2_ChaveCte As String = dr.Item("Doc2_chavecte").ToString ' Chave Cte Eletrônico
                If Len(Trim(EmiDoc2_ChaveCte)) = 44 Then
                    idDoc2Ant = objCTeUtil.idDocAntEle300(EmiDoc2_ChaveCte)
                ElseIf Len(Trim(EmiDoc2_TipoDoc)) = 2 And Len(Trim(EmiDoc2_Numero)) > 0 Then
                    idDoc2Ant = objCTeUtil.idDocAntPap(EmiDoc2_TipoDoc, EmiDoc2_Serie, EmiDoc2_SubSerie, EmiDoc2_Numero, EmiDoc2_DataEmissao)
                End If
                EmiDocAnt = EmiDocAnt & objCTeUtil.EmiDocAnt(EmiDocAnt_CNPJ, EmiDocAnt_CPF, EmiDocAnt_IE_Opc, EmiDocAnt_UF_Opc, EmiDocAnt_xNome, idDoc1Ant, idDoc2Ant)
            Loop
            dr.Close()
            cmd.Dispose()
            '
            '======  Dados do  Dim de Informações do modal Rodoviário==========
            '
            Dim rod As String = ""
            Dim versao As String = "4.00" 'Informar a Versão do Modal
            Dim rod_occ_Opc As String = "" ' Ordens de Coleta associados

            rod = objCTeUtil.rod300(versao, rod_RNTRC, rod_occ_Opc)
            '
            '======  Dados do Dim Ordens de Coleta Associados==========
            '
            Dim occ As String = ""
            'Dim occ_serie_Opc As String = dr.Item("Emitente_cnpj").ToString ' Série da OCC
            'Dim occ_nOcc As String = dr.Item("Emitente_cnpj").ToString ' Número da Ordem de Coleta
            'Dim occ_dEmi As Date = dr.Item("Emitente_cnpj").ToString ' Data de Emissão da Ordem de Coleta
            'Dim occ_CNPJ As String = dr.Item("Emitente_cnpj").ToString ' Número do CNPJ
            'Dim occ_cInt_Opc As String = dr.Item("Emitente_cnpj").ToString ' Código Interno das Tranportadoras
            'Dim occ_IE As String = dr.Item("Emitente_cnpj").ToString ' Inscrição Estadual sem máscara
            'Dim occ_UF As String = dr.Item("Emitente_cnpj").ToString ' sigla da UF
            'Dim occ_fone_Opc As String = dr.Item("Emitente_cnpj").ToString ' número do telefone sem máscara
            '
            'occ = objCTeUtil.occ(occ_serie_Opc, occ_nOcc, occ_dEmi, occ_CNPJ, occ_cInt_Opc, occ_IE, occ_UF, occ_fone_Opc)
            '
            '======  Dados do Dim do Grupo de modal áereo ==========
            '
            Dim aereo As String = ""
            '
            '======  Dados do Dim da informação de manuseio ==========
            '
            Dim cInfManu As String = ""
            '
            '======  Dados do Dim do Grupo de Produtos classificados pela ONU como perigosos==========
            '
            Dim peri As String = ""
            '
            '======  Dados do  Dim de Informações do modal Aquaviário==========
            '
            Dim aquav As String = ""
            '
            '======  Dados do Dim de Identificação da Balsa==========
            '
            Dim balsa As String = ""
            '
            '======  Dados do Dim de Detalhamento dos Containers==========
            '
            Dim detcont As String = ""
            '
            '======  Dados do  Dim de informações dos lacres do container==========
            '
            Dim lacre As String = ""
            '
            '=======Dim do Grupo de Informações das NF das mercadorias transportadas nos Containers do Modal Aquaviário==================
            '
            Dim infNFAquav As String = ""
            '
            '=======Dim do Grupo de Informações das NF-e das mercadorias transportadas nos Containers do Modal Aquaviário==================
            '
            Dim infNFeAquav As String = ""
            '
            '======  Dados do Dim Informações de veículos transportados==========
            '
            Dim veicNovos As String = ""
            '
            '====== Dados do Dim Dados da Cobrança do CT-e==========
            '
            Dim cobr As String = ""
            'Dim cobr_nFat_Opc As String = dr.Item("Emitente_cnpj").ToString ' Número da Fatura
            'Dim cobr_vOrig_Opc As Double = dr.Item("Emitente_cnpj").ToString ' Valor original da fatura (15 posições, sendo 13 inteiras e 2 decimais)
            'Dim cobr_vDesc_Opc As Double = dr.Item("Emitente_cnpj").ToString ' Valor do desconto da fatura (15 posições, sendo 13 inteiras e 2 decimais)
            'Dim cobr_Liq_Opc As Double = dr.Item("Emitente_cnpj").ToString ' Valor líquido da fatura (15 posições, sendo 13 inteiras e 2 decimais)
            'Dim cobr_dup_Opc As String = dr.Item("Emitente_cnpj").ToString ' Dados das duplicatas
            '
            'cobr = objCTeUtil.cobr(cobr_nFat_Opc, cobr_vOrig_Opc, cobr_vDesc_Opc, cobr_Liq_Opc, cobr_dup_Opc)
            '
            '======  Dados do Dim Dados das duplicatas==========
            '
            Dim dup As String = ""
            'Dim dup_nDup As String = dr.Item("Emitente_cnpj").ToString ' Número da Duplicata
            'Dim dup_dVenc As Date = dr.Item("Emitente_cnpj").ToString ' Data de vencimento da duplicata
            'Dim dup_vDup As Double = dr.Item("Emitente_cnpj").ToString ' Valor da duplicata (15 posições, sendo 13 inteiras e 2 decimais)
            '
            'dup = objCTeUtil.dup(dup_nDup, dup_dVenc, dup_vDup)
            '
            '======  Dados do Dim Informações do CT-e de substituição==========
            '
            Dim infCteSub_refNFe As String = ""
            'Dim infCteSub_refNFe_chCte As String = dr.Item("Emitente_cnpj").ToString ' Chave de Acesso do CT-e a ser substituído
            'Dim infCteSub_refNFe_refNFe As String = dr.Item("Emitente_cnpj").ToString ' Chave de Acesso do NF-e emitida pelo Tomador
            'Dim infCteSub_refCTe_indAlteraToma_Opc As String = dr.Item("Emitente_cnpj").ToString ' Informar Indicador de CT-e Alomador
            '
            'infCteSub_refNFe = objCTeUtil.infCteSub_refNFe300(infCteSub_refNFe_chCte, infCteSub_refNFe_refNFe, infCteSub_refCTe_indAlteraToma_Opc)
            '
            '======  Dados do Dim Informações da NF ou CT emitido pelo Tomador==========
            '
            Dim infCteSub_refNF As String = ""
            'Dim infCteSub_refNF_chCte As String = dr.Item("Emitente_cnpj").ToString ' Chave de acesso da NF-e a ser substituído (original)
            'Dim infCteSub_refNF_CNPJ As String = dr.Item("Emitente_cnpj").ToString ' CNPJ do Emitente
            'Dim infCteSub_refNF_mod As String = dr.Item("Emitente_cnpj").ToString ' Modelo do Documento Fiscal
            'Dim infCteSub_refNF_serie As String = dr.Item("Emitente_cnpj").ToString ' Série do documento fiscal
            'Dim infCteSub_refNF_subSerie_Opc As String = dr.Item("Emitente_cnpj").ToString ' Subserie do documento fiscal
            'Dim infCteSub_refNF_nro As String = dr.Item("Emitente_cnpj").ToString ' Número do documento fiscal
            'Dim infCteSub_refNF_valor As Double = dr.Item("Emitente_cnpj").ToString ' Valor do comumento fiscal (15 posições, sendo 13 inteiras e 2 decimais)
            'Dim infCteSub_refNF_dEmi As Date = dr.Item("Emitente_cnpj").ToString ' Data de emissão do documento fiscal
            'infCteSub_refCTe_indAlteraToma_Opc = dr.Item("Emitente_cnpj").ToString ' Informar Indicador de CT-e Alteração de Tomador
            '
            'infCteSub_refNF = objCTeUtil.infCteSub_refNF(infCteSub_refNF_chCte, infCteSub_refNF_CNPJ, infCteSub_refNF_mod, infCteSub_refNF_serie, infCteSub_refNF_subSerie_Opc, infCteSub_refNF_nro, infCteSub_refNF_valor, infCteSub_refNF_dEmi)
            '
            '======  Dados do Dim Informações do CT-e de substituição==========
            '
            Dim infCteSub_refCTe As String = ""
            'Dim infCteSub_refCTe_chCte As String = dr.Item("Emitente_cnpj").ToString  ' Chave de Acesso do CT-e a ser substituído
            'Dim infCteSub_refCTe_refCte As String = dr.Item("Emitente_cnpj").ToString ' Chave de Acesso do CT-e emitida pelo Tomador
            'infCteSub_refCTe_indAlteraToma_Opc = dr.Item("Emitente_cnpj").ToString ' Informar Indicador de CT-e Alteração de Tomador
            '
            'infCteSub_refCTe = objCTeUtil.infCteSub_refCTe300(infCteSub_refCTe_chCte, infCteSub_refCTe_refCte, infCteSub_refCTe_indAlteraToma_Opc)
            '
            '======  Dados do Dim Informações do CT-e de substituição CT-e Anulação==========
            '
            Dim infCteSub_refCTeAnu As String = ""
            'Dim infCteSub_refCTeAnu_chCte As String = dr.Item("Emitente_cnpj").ToString ' Chave de Acesso do CT-e a ser substituído
            'Dim infCteSub_refCTeAnu_refCteAnu As String = dr.Item("Emitente_cnpj").ToString ' Chave de Acesso do CT-e a ser substituído
            'infCteSub_refCTe_indAlteraToma_Opc = dr.Item("Emitente_cnpj").ToString 'Informar Indicador de CT-e Alteração de Tomador
            '
            'infCteSub_refCTeAnu = objCTeUtil.infCteSub_refCTeAnu300(infCteSub_refCTeAnu_chCte, infCteSub_refCTeAnu_refCteAnu, infCteSub_refCTe_indAlteraToma_Opc)
            '
            '======  Dados do Dim informações do CTe Multimodal==========
            '
            Dim infCTeMultimodal As String = ""
            'infCTeMultimodal = objCTeUtil.infCTeMultimodal(infCTeMultimodal_chCte)
            '
            '======  Dados do Dim Detalhamento do CT-e complementado==========
            '
            Dim infCteComp As String = ""
            If identificador_tpCTe = 1 Then
                cmd.CommandText = "Select * from Notas_fiscais_nfes_referenciados (nolock) where nota_fiscal = " & lnota_fiscal
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    Dim infCte_chave As String = dr.Item("Chave_nfe").ToString ' Número da Chave de Acesso das NF-e
                    '
                    infCteComp = infCteComp + objCTeUtil.infCteComp300(infCte_chave)
                Loop
                dr.Close()
                cmd.Dispose()
            End If
            '
            '
            '======  Dados do Dim do Grupo de Detalhamento do CT-e do tipo Anulação de Valores==========
            '
            Dim infCteAnu As String = ""
            'Dim infCteAnu_chCte As String = dr.Item("Emitente_cnpj").ToString ' Número da Chave de Acesso do Ct-e original a ser anulado ou substituído
            'Dim infCteAnu_dEmi As Date = dr.Item("Emitente_cnpj").ToString ' Data de emissão da declaração do tomador não contribuinte do ICMS
            '
            'infCteAnu = objCTeUtil.infCteAnu(infCteAnu_chCte, infCteAnu_dEmi)
            '
            '======  Dim InfCorrecao ==========
            '
            Dim autXML As String = ""
            'Dim autXML_CNPJ As String = dr.Item("Emitente_cnpj").ToString ' informar CNPJ
            'Dim autXML_CPF As String = dr.Item("Emitente_cnpj").ToString ' ou CPF
            'autXML = autXML + objCTeUtil.autXML(autXML_CNPJ, autXML_CPF)
            '
            '======  Dados do Dim do Grupo de Corte de Voo ==========
            '
            Dim infRespTec As String = ""
            'Dim infRespTec_CNPJ As String = dr.Item("Emitente_cnpj").ToString ' informar o CNPJ da PJ responsável técnica pela emissão do documento fiscal eletrônico
            'Dim infRespTec_xContato As String = dr.Item("Emitente_cnpj").ToString ' informar o nome da pessoa de contato
            'Dim infRespTec_email As String = dr.Item("Emitente_cnpj").ToString ' informar o e-mail da PJ a ser contatada
            'Dim infRespTec_fone As String = dr.Item("Emitente_cnpj").ToString ' informar o telefone da PJ a ser contatada
            'Dim infRespTec_idCSRT As String = dr.Item("Emitente_cnpj").ToString ' informar o identificador do código de segurança do responsavel técnico
            'Dim infRespTec_hashCSRT As String = dr.Item("Emitente_cnpj").ToString ' inforamr o hash do token do código de segurança do responsavel técnico
            '
            'infRespTec = objCTeUtil.infRespTec(infRespTec_CNPJ, infRespTec_xContato, infRespTec_email, infRespTec_fone, infRespTec_idCSRT, infRespTec_hashCSRT)
            '
            '======  Dados do Dim do Grupo de Corte de Voo ==========
            '
            '
            Dim infRespTec2 As String = ""
            '
            '======  Dados do Dim do Grupo de Corte de Voo ==========
            '
            Dim infCTeSupl As String = ""
            Dim infCTeSupl_URL As String = Função_Seleciona_url_qrcode(emi_UF, identificador_tpAmb)
            Dim infCTeSupl_chaveCTe As String = ChaveCte
            Dim infCTeSupl_tpAmb As Long = identificador_tpAmb
            Dim infCTeSupl_QRCode As String = ""
            Dim resultado As Long

            infCTeSupl = objCTeUtil.infCTeSupl(infCTeSupl_URL, infCTeSupl_chaveCTe, infCTeSupl_tpAmb, infCTeSupl_nomeCertificado, infCTeSupl_QRCode, resultado, msgResultado)
            '
            '======  Dados do  Conhecimento de Tranporte Eletrônico==========
            '
            Dim CTe As String
            'Dim versao As String = dr.Item("Emitente_cnpj").ToString
            '
            '===================Grupo de Informações do CT-e Normal e Substituto=======================
            '
            Dim infCTeNorm As String = ""
            '
            '===================Grupo de Informações do CT-e Normal e Substituto=======================
            '
            Dim infCTeNorm_infDoc_Opc As String = infNF + infNFe + infOutros     ' Informaçoes dos documentos fiscais que acobertam a carga
            Dim infCTeNorm_infGlobalizado_Opc As String = ""       ' Informações do CT-e globalizado, preencher com informações adicionais, legislação do regime especial, etc
            '
            If identificador_tpCTe = 1 Then 'CTe Complementar
                infCTeNorm = infCteComp
            Else
                infCTeNorm = objCTeUtil.infCTeNorm300(infCarga, infCTeNorm_infDoc_Opc, EmiDocAnt, rod, veicNovos, cobr, infCteSub_refNFe, infCTeNorm_infGlobalizado_Opc, infCTeMultimodal)
            End If
            CTe = objCTeUtil.CTe_v3a(versao, ChaveCte, identificador, compl, emi, reme, exped, receb, dest, vPrest, imp, infCTeNorm, autXML, infRespTec, infCTeSupl)
            '
            '===================grupo de Observações gerais do Contribuinte=======================
            '
            'obsCont_xCampo = ""                 ' Identificação do campo
            obsCont_xTexto = ""                 ' Conteúdo do campo
            '
            'obsCont = objCTeUtil.obsCont(obsCont_xCampo, obsCont_xTexto)
            Dim larquivo_xml As String

            larquivo_xml = My.Application.Info.DirectoryPath & "\Cte\" & emi_CNPJ & "\NaoAssinada\" & ChaveCte & ".xml"
            If IO.File.Exists(larquivo_xml) Then System.IO.File.Delete(larquivo_xml)
            Dim objStream1 As New System.IO.FileStream(larquivo_xml, IO.FileMode.Create)
            Dim Arq1 As New System.IO.StreamWriter(objStream1)
            Arq1.Write(CTe)
            Arq1.Close()

            '
            'Assina e Envia o Xml para o Sefaz
            '
            Dim CteAssinado As String = ""
            Dim cStat As Long
            Dim nroRecibo As String = ""
            Dim dhRecibo As String = ""
            '
            ' Cte         - informar com a CT-e a ser transmitida, não é necessário validar nem assinar
            '
            ' CteAssinada - é devolvido pela DLL se a chamada for realizada com sucesso.
            '
            ' NroRecibo   - é devolvido pela DLL se a CT-e for transmitida corretamente,
            '               este número é necessário para buscar o resultado do processamento da CT-e
            '               ========== o NroRecibo não indica que a CT-e foi autorizada ==============
            Dim msgDados As String = ""
            Dim msgRetWS As String = ""
            Dim tMed As String = ""
            Dim nroProtocolo As String  ' número do protocolo de autorização de uso da NF-e, este número é necessário para cancelar a NF-e
            Dim dhProtocolo As String   ' data e hora de autorização de uso da NF-e
            Dim procCte As String = ""
            Do

                'cStat = objCTeUtil.EnviaCTe(emi_UF, CTe, nroRecibo, infCTeSupl_nomeCertificado, versao, msgDados, msgRetWS, msgResultado, CteAssinado, gcte_proxy, gcte_proxy_usuario, gcte_proxy_senha, licenca)
                procCte = objCTeUtil.EnviaCTeSinc(emi_UF, versao, infCTeSupl_nomeCertificado, CTe, msgDados, msgRetWS, cStat, msgResultado, CteAssinado, nroProtocolo, dhProtocolo, gcte_proxy, gcte_proxy_usuario, gcte_proxy_senha, licenca)
                '
                '  se cStat = 105 - Lote em processamento, significa que a SEFAZ ainda não conseguiu processar a NF-e e a aplicação deve persistir até que o cStat seja diferent de 105
                '
                If cStat <> 105 Then
                    Exit Do
                End If
            Loop
            ' análise do retorno da chamada da função enviaNFeSCAN
            '
            ' resultado:
            '
            ' 5000-7000 ------> falha na chamda da função, vide a mensagem de erro e corrija o problema - http://www.flexdocs.com.br/guiaNFe/WS.NFe.enviaNFe2G.html
            '
            ' 103 -------> SUCESSO! - LOTE RECEBIDO pelo WS, guarde o NroRecibo e efetua a busca do resultado do processamento
            '
            '              IMPORTANTE, ter o número do recibo não significa que a NF-e está autorizada, ainda é necessário buscar o resultado do processamento
            '
            ' 108/109 ---> WS com problemas, sem condições de receber lotes
            ' 2xx -------> FALHA - existe algum problema com o lote enviado, verifique o código do erro e corrigir o erro
            '
            Dim lcomentarios As String = ""
            Dim lstatus As String = ""
            Select Case cStat
                Case 203 To 6423                ' erro da DLL
                    lstatus = "EI"
                    lcomentarios = "EnviaCTeSinc " & cStat & msgResultado
                    'Rejeição: Duplicidade de NF-e [nRec:
                    If Mid(lcomentarios, 1, 35) = "EnviaCTeSinc 204Rejeição: Duplicidade" Then
                        lstatus = "AS"
                        nroProtocolo = Mid(lcomentarios, 52, 16)
                        dhProtocolo = Mid(lcomentarios, 77, 25)
                    End If
                    Dim sqlBuilder As New System.Text.StringBuilder
                    sqlBuilder.Append("update notas_fiscais set Xml = '" & CTe & "' where nota_fiscal = " & lnota_fiscal)
                    BDexecuta_query(sqlBuilder)
                Case 100    ' CT-e autorizada
                    '
                    'Cria o arquivo XML Assinado
                    '
                    larquivo_xml = My.Application.Info.DirectoryPath & "\Cte\" & emi_CNPJ & "\" & ChaveCte & ".xml"

                    If IO.File.Exists(larquivo_xml) Then System.IO.File.Delete(larquivo_xml)
                    Dim objStream As New System.IO.FileStream(larquivo_xml, IO.FileMode.Create)
                    Dim Arq As New System.IO.StreamWriter(objStream)
                    Arq.Write(procCte)
                    Arq.Close()

                    lstatus = "AS"
                    lcomentarios = "EnviaCTeSinc " & "Autorizado Prot. " & nroProtocolo
                    Dim sqlBuilder As New System.Text.StringBuilder
                    sqlBuilder.Append("update notas_fiscais set Xml = '" & procCte & "',Protocolo_autorizacao = '" & nroProtocolo & "',Data_hora_autorizacao='" & dhProtocolo & "' where nota_fiscal = " & lnota_fiscal)
                    BDexecuta_query(sqlBuilder)
                Case 101    ' NF-e denegada
                    lstatus = "RS"
                    lcomentarios = "EnviaCTeSinc " & "NF Denegada Prot " & nroProtocolo & " Msg: " & msgResultado
                    Dim sqlBuilder As New System.Text.StringBuilder
                    sqlBuilder.Append("update notas_fiscais set Xml = '" & CTe & "' where nota_fiscal = " & lnota_fiscal)
                    BDexecuta_query(sqlBuilder)

                Case 106    ' Lote não localizado, verifique se o número do recibdo está correto
                    lstatus = "EI"
                    lcomentarios = "EnviaCTeSinc Lote não localizado"
                    Dim sqlBuilder As New System.Text.StringBuilder
                    sqlBuilder.Append("update notas_fiscais set Xml = '" & CTe & "' where nota_fiscal = " & lnota_fiscal)
                    BDexecuta_query(sqlBuilder)

                Case 108, 109 'Problema na recepção dos webservices do sefaz
                    lstatus = "PS"
                    lcomentarios = "EnviaCTeSinc Prob. Sefaz "
                    Dim sqlBuilder As New System.Text.StringBuilder
                    sqlBuilder.Append("update notas_fiscais set Xml = '" & CTe & "' where nota_fiscal = " & lnota_fiscal)
                    BDexecuta_query(sqlBuilder)

                Case Else
                    lstatus = "EI"
                    lcomentarios = "EnviaCTeSinc " & cStat & msgResultado & msgRetWS
                    Dim sqlBuilder As New System.Text.StringBuilder
                    sqlBuilder.Append("update notas_fiscais set Xml = '" & CTe & "' where nota_fiscal = " & lnota_fiscal)
                    BDexecuta_query(sqlBuilder)
            End Select
            Call BDInclui_evento_nf(lcomentarios, lnota_fiscal, lstatus)
            '
            'Versão 3.0
            '
            'Dim lstatus As String = "PS"
            'Select Case cStat
            'Case 5505
            'lstatus = "EE"
            'lcomentarios = "EnviaCTe" & "/" & msgResultado
            'Case 5000 To 7003                ' problemas no envio da DLL para o WS
            'lcomentarios = "EnviaCTe" & "/" & msgResultado
            'Case 103 'Deu certo
            'lcomentarios = "A Ct-e foi enviada para o Sefaz"
            'lstatus = "ES"
            'cmd.CommandText = "update notas_fiscais set numero_recibo = '" & nroRecibo & "' where nota_fiscal = " & lnota_fiscal
            'cmd.ExecuteNonQuery()
            'Case 108, 109 'problemas na recepção do WS
            'lcomentarios = "EnviaCTe" & msgResultado & "/" & cStat & "/" & versao & "/" & msgDados & "/" & msgRetWS
            'Case Else 'Outros Erros'
            'lcomentarios = "EnviaCTe" & msgResultado & "/" & cStat & "/" & versao & "/" & msgDados & "/" & msgRetWS
            'End Select
            'Call BDInclui_evento_nf(lcomentarios, lnota_fiscal, lstatus)
            objCTeUtil = Nothing

            conn.Close()
            conn.Dispose()
        End If
    End Sub
    Public Sub BDInclui_evento_nf(ByVal lcomentario As String, ByVal lnota_fiscal As Long, ByVal lstatus As String)
        Dim ldata_hota As String
        ldata_hota = String.Format("{0:yyyyMMdd HH:mm:ss}", Now)
        lcomentario = lcomentario.Replace("http://www.portalfiscal.inf.br/nfe", "")
        lcomentario = lcomentario.Replace("'", "#")
        '
        'Inclui evento nf
        '
        Dim lNota_fiscal_evento As Long
        lNota_fiscal_evento = BDProximo_codigo("Notas_fiscais_eventos", "Nota_fiscal_evento", "") 'Gera do próximo código do romaneio
        Dim sqlBuilder As New System.Text.StringBuilder
        With sqlBuilder 'mesma coisa que o sSql que é uma variável string
            .Append("insert into Notas_fiscais_eventos ")
            .Append(" (Nota_fiscal_evento,")
            .Append("Nota_fiscal,Data_hora,")
            .Append("Status,Comentarios,Usuario) ")
            .Append(" values (")
            .Append(lNota_fiscal_evento & ",")
            .Append(lnota_fiscal & ",'" & ldata_hota & "','")
            .Append(lstatus & "','" & Left(lcomentario, 5000) & "','NFEService')")
        End With
        BDexecuta_query(sqlBuilder)
        sqlBuilder.Remove(0, sqlBuilder.Length)
        '
        'Atualiza o Status da Nota Fiscal
        '
        'Pega o Nome da Máquina
        Dim lComputerName As String
        lComputerName = System.Net.Dns.GetHostName
        With sqlBuilder 'mesma coisa que o sSql que é uma variável string
            .Append("Update notas_fiscais ")
            .Append("Set ")
            .Append("Status = '" & lstatus & "',")
            .Append("Status_comentarios = '" & Microsoft.VisualBasic.Left(lcomentario, 200) & "',")
            .Append("Status_ultimo_evento = '" & ldata_hota & "' ")
            .Append("where nota_fiscal = " & lnota_fiscal)
        End With
        BDexecuta_query(sqlBuilder)
    End Sub

    Public Function BDexecuta_query(ByRef ComandoSql As System.Text.StringBuilder)
        BDexecuta_query = False
        Try
            Dim conn As DbConnection = Me.dbfactory.CreateConnection
            conn.ConnectionString = My.Settings.ConnectionString
            Dim cmd As DbCommand = conn.CreateCommand
            conn.Open()
            cmd.CommandText = ComandoSql.ToString
            cmd.ExecuteNonQuery()
            conn.Close()
            cmd.Dispose()
            conn.Dispose()
            BDexecuta_query = True
        Catch err As Exception
            'MessageBox.Show(err.Message)
        End Try
    End Function

    Public Function BDProximo_codigo(ByVal tabela As String, ByRef Campo As String, ByRef condicao As String) As Long
        BDProximo_codigo = 0
        Try
            Dim conn As DbConnection = Me.dbfactory.CreateConnection
            conn.ConnectionString = My.Settings.ConnectionString
            Dim cmd As DbCommand = conn.CreateCommand
            Dim dr As DbDataReader
            BDProximo_codigo = 0
            conn.Open()
            cmd.CommandText = "Select max(" & Campo & ") as proximo_codigo from " & tabela & " (nolock)"
            If condicao <> "" Then
                cmd.CommandText = cmd.CommandText & " where " & condicao
            End If
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                BDProximo_codigo = CInt(0 & dr.Item("proximo_codigo").ToString) + 1
                dr.Close()
                conn.Close()
                cmd.Dispose()
                conn.Dispose()
            End If
        Catch err As Exception
            'MessageBox.Show(err.Message)
        End Try
    End Function

    Public Function Retira_esp_pont_hifem_colc_bar_ace_vir_igu(ByVal ltexto As String, ByVal lespaco As Boolean, ByVal lponto As Boolean, ByVal lhifem As Boolean, ByVal lcolchete As Boolean, ByVal lbarra As Boolean, ByVal lacentos As Boolean, ByVal lvirgula As Boolean, ByVal ligual As Boolean) As String

        If lespaco Then ltexto = ltexto.Replace(" ", "")
        If lponto Then ltexto = ltexto.Replace("-", "")
        If lhifem Then ltexto = ltexto.Replace(".", "")
        If lcolchete Then
            ltexto = ltexto.Replace("(", "")
            ltexto = ltexto.Replace(")", "")
        End If
        If lbarra Then
            ltexto = ltexto.Replace("/", "")
            ltexto = ltexto.Replace("\", "")
        End If
        If lacentos Then
            ltexto = ltexto.Replace("Ã", "A")
            ltexto = ltexto.Replace("Õ", "O")
            ltexto = ltexto.Replace("ã", "a")
            ltexto = ltexto.Replace("õ", "o")
            ltexto = ltexto.Replace("Â", "A")
            ltexto = ltexto.Replace("Ê", "E")
            ltexto = ltexto.Replace("Ô", "O")
            ltexto = ltexto.Replace("Û", "U")
            ltexto = ltexto.Replace("â", "a")
            ltexto = ltexto.Replace("ê", "e")
            ltexto = ltexto.Replace("ô", "o")
            ltexto = ltexto.Replace("û", "i")
            ltexto = ltexto.Replace("Á", "A")
            ltexto = ltexto.Replace("É", "E")
            ltexto = ltexto.Replace("Í", "I")
            ltexto = ltexto.Replace("Ó", "O")
            ltexto = ltexto.Replace("Ú", "U")
            ltexto = ltexto.Replace("á", "a")
            ltexto = ltexto.Replace("é", "e")
            ltexto = ltexto.Replace("í", "i")
            ltexto = ltexto.Replace("ó", "o")
            ltexto = ltexto.Replace("ú", "u")
            ltexto = ltexto.Replace("ç", "c")
            ltexto = ltexto.Replace("Ç", "C")
        End If
        If lvirgula Then ltexto = ltexto.Replace(",", "")
        If ligual Then ltexto = ltexto.Replace("=", "")
        Retira_esp_pont_hifem_colc_bar_ace_vir_igu = ltexto
    End Function

    Public Function BDretorna_campo(ByVal tabela As String, ByRef Campo As String, ByRef condicao As String) As String
        BDretorna_campo = Nothing
        Try
            Dim conn As DbConnection = dbfactory.CreateConnection
            conn.ConnectionString = My.Settings.ConnectionString
            Dim cmd As DbCommand = conn.CreateCommand
            Dim dr As DbDataReader
            conn.Open()
            cmd.CommandText = "Select " & Campo & " from " & tabela & " where " & condicao
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                BDretorna_campo = dr.Item(0).ToString
                dr.Close()
                conn.Close()
                cmd.Dispose()
                conn.Dispose()
            End If
        Catch err As Exception
            'MessageBox.Show(err.Message)
        End Try
    End Function

    Public Sub NfeUtil_Consulta_Cte_Sefaz(ByVal lcertificado As String, ByVal lChaveNFe As String, ByVal lversao As String, ByVal lsiglaWS As String, ByVal lnumero_recibo As String, ByVal lnota_fiscal As Long, ByVal ltipo_ambiente As Integer, ByVal lEmitente_cnpj As String, ByVal lmodo As String, ByVal chaveCTe As String, ByVal siglaUF As String)
        'Dim cteproc As String
        Dim cStat As Long
        '
        '  Consulta Status da NF-e
        '
        '
        Dim msgDados As String
        Dim msgRetWS As String
        Dim msgResultado As String
        '
        '  As variáveis do proxy devem ser informadas se necessário
        '
        '  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
        '
        Dim proxy As String
        Dim usuario As String
        Dim senha As String
        '
        '  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
        '
        proxy = gcte_proxy
        usuario = gcte_proxy_usuario
        senha = gcte_proxy_senha
        msgDados = ""
        msgRetWS = ""
        msgResultado = ""
        Dim emi_CNPJ As String = ""
        emi_CNPJ = Microsoft.VisualBasic.Left(lEmitente_cnpj, 2) + Microsoft.VisualBasic.Mid(lEmitente_cnpj, 4, 3) + Microsoft.VisualBasic.Mid(lEmitente_cnpj, 8, 3) + Microsoft.VisualBasic.Mid(lEmitente_cnpj, 12, 4) + Microsoft.VisualBasic.Right(lEmitente_cnpj, 2) ' CNPJ do emitente sem máscara de formatação
        Dim Resultado As Long

        Dim objCTeUtil As Object

        objCTeUtil = CreateObject("CTe_Util.Util")

        '
        ' chave da licenca de uso da DLL
        '
        Dim licenca As String = BDretorna_campo("Filiais", "Cte_chave_flexdocs", "cnpj = '" & lEmitente_cnpj & "'")
        '
        ' define as variáveis que passam/recebem informações importantes
        '
        Dim CteAssinada As String = ""
        Dim procCte As String = ""       ' procNFe -> NF-e + protocolo de autorização de uso da NF-e, deve ser mantido em arquivo e distribuído ao destinatário.
        '
        ' parâmetros novos
        '
        Dim nroProtocolo As String  ' número do protocolo de autorização de uso da NF-e, este número é necessário para cancelar a NF-e
        Dim dhProtocolo As String   ' data e hora de autorização de uso da NF-e
        Dim cMsg As String          ' código da mensagem da SEFAZ, a SEFAZ pode utiliza-lo como canal de comunicação com o emissor
        Dim xMsg As String          ' literal da mensagem da SEFAZ
        '
        '
        '  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
        '
        '
        'cteproc = ""
        nroProtocolo = ""
        dhProtocolo = ""
        cMsg = ""
        xMsg = ""

        '
        '  carregar arquivo XML da NF-e assinada na string NFeAssinada, a NF-e assinada
        '  é necessário para montar o procNFe
        '
        On Error Resume Next
        '
        'Abre o arquivo XML para consultar no sefaz
        '
        Dim larquivo_xml As String
        larquivo_xml = My.Application.Info.DirectoryPath & "\cte\" & emi_CNPJ & "\" & lChaveNFe & ".xml"

        Dim objStream As New System.IO.FileStream(larquivo_xml, IO.FileMode.Open)
        Dim Arq As New System.IO.StreamReader(objStream)
        CteAssinada = Arq.ReadLine
        Arq.Close()
        '
        'Processa a busca da NFE no Sefaz
        '
        Do
            cStat = objCTeUtil.BuscaCTe(lsiglaWS, ltipo_ambiente, siglaUF, CteAssinada, lnumero_recibo, procCte, nroProtocolo, dhProtocolo, lcertificado, lversao, proxy, cMsg, msgResultado, proxy, usuario, senha, licenca)
            '
            '  se cStat = 105 - Lote em processamento, significa que a SEFAZ ainda não conseguiu processar a NF-e e a aplicação deve persistir até que o cStat seja diferent de 105
            '
            If cStat <> 105 Then
                Exit Do
            End If
        Loop

        '
        '   tratar o resultado da chamada:
        '
        '
        '           WS chamada com sucesso
        '
        '           105 – lote em processamento -> tentar novamente
        '           106 – lote não localizado   -> tentar enviar o lote novamente ou verificar se o nroRecibo está correto
        '           100 – NF-e autorizada       -> OK
        '           2xx – motivo de rejeição do WS -> erro na elaboração da NF-e, verificar o código de erro e corrigir a NF-e
        Dim lcomentarios As String = ""
        Dim lstatus As String = ""
        Select Case cStat
            Case 203 To 6423                ' erro da DLL
                lstatus = "EI"
                lcomentarios = "BuscaCte2G " & cStat & msgResultado
                'Rejeição: Duplicidade de NF-e [nRec:
                If Mid(lcomentarios, 1, 35) = "BuscaCte2G 204Rejeição: Duplicidade" Then
                    lstatus = "AS"
                    nroProtocolo = Mid(lcomentarios, 52, 16)
                    dhProtocolo = Mid(lcomentarios, 77, 25)
                End If
            Case 100    ' NF-e autorizada
                If IO.File.Exists(larquivo_xml) Then System.IO.File.Delete(larquivo_xml)
                Dim objStream1 As New System.IO.FileStream(larquivo_xml, IO.FileMode.Create)
                Dim Arq1 As New System.IO.StreamWriter(objStream1)
                Arq1.Write(procCte)
                Arq1.Close()
                If lmodo <> "RX" Then
                    lstatus = "AS"
                    lcomentarios = "BuscaCte2G " & "Autorizado Prot. " & nroProtocolo
                Else
                    lstatus = "RX"
                    lcomentarios = "BuscaCte2G Regerado o XML"
                    Call BDInclui_evento_nf(lcomentarios, lnota_fiscal, lstatus)
                    lstatus = "FN"
                    lcomentarios = "BuscaCte2G Finalizado"
                    Call BDInclui_evento_nf(lcomentarios, lnota_fiscal, lstatus)
                    Exit Sub
                End If
            Case 101    ' NF-e denegada
                lstatus = "RS"
                lcomentarios = "BuscaCte2G " & "NF Denegada Prot " & nroProtocolo & " Msg: " & msgResultado
            Case 106    ' Lote não localizado, verifique se o número do recibdo está correto
                lstatus = "EI"
                lcomentarios = "BuscaCte2G Lote não localizado"
            Case 108, 109 'Problema na recepção dos webservices do sefaz
                lstatus = "PS"
                lcomentarios = "BuscaCte2G Prob. Sefaz "
            Case Else
                lstatus = "EI"
                lcomentarios = "BuscaCte2G " & cStat & msgResultado & cMsg
        End Select
        Call BDInclui_evento_nf(lcomentarios, lnota_fiscal, lstatus)
        If lstatus = "AS" Then
            Dim sqlBuilder As New System.Text.StringBuilder
            sqlBuilder.Append("update notas_fiscais set Protocolo_autorizacao = '" & nroProtocolo & "',Data_hora_autorizacao='" & dhProtocolo & "' where nota_fiscal = " & lnota_fiscal)
            BDexecuta_query(sqlBuilder)
        Else
            If lmodo = "RX" Then
                lstatus = "FN"
                lcomentarios = "BuscaCte2G Finalizado"
                Call BDInclui_evento_nf(lcomentarios, lnota_fiscal, lstatus)
            End If
        End If

        ' libera classe
        '
        objCTeUtil = Nothing
    End Sub
    Private Function Processo_Distribuir(ByVal lnota_fiscal As Long, ByVal lfilial As Integer, ByVal lfornecedor_cliente As Long, ByVal lchave_acesso As String) As Boolean
        Dim conn As DbConnection = Me.dbfactory.CreateConnection
        conn.ConnectionString = My.Settings.ConnectionString
        Dim cmd As DbCommand = conn.CreateCommand
        Dim dr As DbDataReader

        Dim lFilialNfe_email_servidor As String = ""
        Dim lFilialNfe_email_origem_envio As String = ""
        Dim lFilialNfe_email_copia_pdf As String = ""
        Dim lFilialNfe_email_assunto As String = ""
        Dim lFilialNfe_email_corpo As String = ""
        Dim lFilialNfe_email_copia_xml As String = ""
        Dim lFilialNfe_senha_servidor_email As String = ""

        Dim lFCNfe_email_copia_pdf As String = ""
        Dim lFCNfe_email_copia_xml As String = ""
        Dim lFCNfe_email_copia_pdf_xml As String = ""
        Dim lemitente_cnpj As String = ""
        Dim lchavenfe As String = ""
        Dim lEmail_alternativo As String = ""

        Try
            conn.Open()
            cmd.CommandText = "Select Nfe_senha_servidor_email,Nfe_email_copia_xml,Servidor_email,Nfe_email_origem_envio,Nfe_email_copia_pdf,Nfe_email_assunto,Nfe_email_corpo FROM Filiais where filial = " & lfilial
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                lFilialNfe_email_servidor = dr.Item("Servidor_email").ToString
                lFilialNfe_email_origem_envio = dr.Item("Nfe_email_origem_envio").ToString
                lFilialNfe_email_copia_pdf = dr.Item("Nfe_email_copia_pdf").ToString
                lFilialNfe_email_assunto = dr.Item("Nfe_email_assunto").ToString
                lFilialNfe_email_corpo = dr.Item("Nfe_email_corpo").ToString
                lFilialNfe_email_copia_xml = dr.Item("Nfe_email_copia_xml").ToString
                lFilialNfe_senha_servidor_email = dr.Item("Nfe_senha_servidor_email").ToString
            End If
            dr.Close()
            If lFilialNfe_email_servidor = "" Or lFilialNfe_email_origem_envio = "" Or lFilialNfe_senha_servidor_email = "" Then
                Processo_Distribuir = False
                Exit Function
            End If
            cmd.CommandText = "Select Nfe_email_copia_pdf,Nfe_email_copia_xml,Nfe_email_copia_pdf_xml,cnpj FROM Fornecedores_clientes where fornecedor_cliente = " & lfornecedor_cliente
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                lFCNfe_email_copia_pdf = dr.Item("Nfe_email_copia_pdf").ToString
                lFCNfe_email_copia_xml = dr.Item("Nfe_email_copia_xml").ToString
                lFCNfe_email_copia_pdf_xml = dr.Item("Nfe_email_copia_pdf_xml").ToString
            End If
            dr.Close()
            cmd.CommandText = "Select Chave_acesso,emitente_cnpj,Email_alternativo FROM Notas_fiscais where nota_fiscal = " & lnota_fiscal
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                lemitente_cnpj = Microsoft.VisualBasic.Left(dr.Item("Emitente_cnpj").ToString, 2) + Microsoft.VisualBasic.Mid(dr.Item("Emitente_cnpj").ToString, 4, 3) + Microsoft.VisualBasic.Mid(dr.Item("Emitente_cnpj").ToString, 8, 3) + Microsoft.VisualBasic.Mid(dr.Item("Emitente_cnpj").ToString, 12, 4) + Microsoft.VisualBasic.Right(dr.Item("Emitente_cnpj").ToString, 2) ' CNPJ do emitente sem máscara de formatação
                lchavenfe = dr.Item("chave_acesso").ToString
                lEmail_alternativo = dr.Item("Email_alternativo").ToString
            End If
            dr.Close()
            conn.Close()
            cmd.Dispose()
            conn.Dispose()

            Processo_Distribuir = False
            '
            'Primeiro Envia o PDF
            '
            If lFCNfe_email_copia_pdf <> "" Or lEmail_alternativo <> "" Or lFilialNfe_email_copia_pdf <> "" Then
                Dim mMailMessagePdf As New MailMessage()
                mMailMessagePdf.From = New MailAddress(lFilialNfe_email_origem_envio)
                If lEmail_alternativo <> "" Then
                    mMailMessagePdf.To.Add(New MailAddress(lEmail_alternativo))
                Else
                    If lFCNfe_email_copia_pdf <> "" Then mMailMessagePdf.To.Add(New MailAddress(lFCNfe_email_copia_pdf))
                    If lFilialNfe_email_copia_pdf <> "" Then mMailMessagePdf.CC.Add(New MailAddress(lFilialNfe_email_copia_pdf))
                End If
                mMailMessagePdf.Subject = lFilialNfe_email_assunto
                mMailMessagePdf.Body = lFilialNfe_email_corpo
                mMailMessagePdf.IsBodyHtml = True
                mMailMessagePdf.Priority = MailPriority.High
                mMailMessagePdf.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess
                Dim larquivopdf As String = My.Application.Info.DirectoryPath & "\Nfe\" & lemitente_cnpj & "\" & lchavenfe & ".pdf"
                If IO.File.Exists(larquivopdf) Then
                    mMailMessagePdf.Attachments.Add(New Attachment(larquivopdf))
                End If

                Dim mSmtpClientPdf As New SmtpClient(lFilialNfe_email_servidor)
                Dim statusPdf As New System.Net.Mail.SmtpStatusCode
                mSmtpClientPdf.Credentials = New Net.NetworkCredential(lFilialNfe_email_origem_envio, lFilialNfe_senha_servidor_email)
                mSmtpClientPdf.Port = 587
                mSmtpClientPdf.DeliveryMethod = SmtpDeliveryMethod.Network
                mSmtpClientPdf.Timeout = 50000
                If lFilialNfe_email_servidor = "smtp.live.com" Then mSmtpClientPdf.EnableSsl = True
                mSmtpClientPdf.Send(mMailMessagePdf)

                mMailMessagePdf.Dispose()
                mMailMessagePdf = Nothing
                mSmtpClientPdf = Nothing
                Processo_Distribuir = True
            End If
            '
            'Segundo Envia o XML
            '
            If lFCNfe_email_copia_xml <> "" Or lEmail_alternativo <> "" Or lFilialNfe_email_copia_xml <> "" Then
                Dim mMailMessageXml As New MailMessage()

                mMailMessageXml.From = New MailAddress(lFilialNfe_email_origem_envio)
                If lEmail_alternativo <> "" Then
                    mMailMessageXml.To.Add(New MailAddress(lEmail_alternativo))
                Else
                    If lFCNfe_email_copia_xml <> "" Then mMailMessageXml.To.Add(New MailAddress(lFCNfe_email_copia_xml))
                    If lFilialNfe_email_copia_xml <> "" Then mMailMessageXml.CC.Add(New MailAddress(lFilialNfe_email_copia_xml))
                End If
                mMailMessageXml.Subject = lFilialNfe_email_assunto
                mMailMessageXml.Body = lFilialNfe_email_corpo
                mMailMessageXml.IsBodyHtml = True
                mMailMessageXml.Priority = MailPriority.High
                mMailMessageXml.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess
                Dim larquivoxml As String = My.Application.Info.DirectoryPath & "\Nfe\" & lemitente_cnpj & "\" & lchavenfe & ".xml"
                If IO.File.Exists(larquivoxml) Then
                    mMailMessageXml.Attachments.Add(New Attachment(larquivoxml))
                End If

                Dim mSmtpClientXml As New SmtpClient(lFilialNfe_email_servidor)
                Dim statusXml As New System.Net.Mail.SmtpStatusCode
                mSmtpClientXml.Port = 587
                mSmtpClientXml.DeliveryMethod = SmtpDeliveryMethod.Network
                mSmtpClientXml.Timeout = 50000
                If lFilialNfe_email_servidor = "smtp.live.com" Then mSmtpClientXml.EnableSsl = True
                mSmtpClientXml.Credentials = New Net.NetworkCredential(lFilialNfe_email_origem_envio, lFilialNfe_senha_servidor_email)
                mSmtpClientXml.Send(mMailMessageXml)

                mMailMessageXml.Dispose()
                mMailMessageXml = Nothing
                mSmtpClientXml = Nothing
                Processo_Distribuir = True
            End If
            If lEmail_alternativo <> "" Then
                '
                'Limpa Email Alternativo
                '
                Dim sqlBuilder As New System.Text.StringBuilder
                sqlBuilder.Append("update notas_fiscais set Email_alternativo = '' where nota_fiscal = " & lnota_fiscal)
                BDexecuta_query(sqlBuilder)
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString)
            Processo_Distribuir = False
        End Try
    End Function
    Private Sub NfeUtil_Cancela_cte(ByVal lnota_fiscal As Long)
        Dim conn As DbConnection = Me.dbfactory.CreateConnection
        conn.ConnectionString = My.Settings.ConnectionString
        Dim cmdNf As DbCommand = conn.CreateCommand
        Dim dr As DbDataReader
        '
        '  Cancelamento da NF-e
        '
        '  Esta funcionaliade deve ser utilizada para cancelar
        '  uma NF-e autorizada e ainda não tenha ocorrido o fato
        '  gerador (circulação da mercadoria).
        '  Ex. falta de mercadoria, divergência de quantidade, preço, etc.
        '  desistência do comprador, etc.
        '
        '  veja detalhes da funcionalidade em: http://www.flexdocs.com.br/guiaNFe/WS.canc.cancelaNF2G.html
        '
        Dim msgDados As String
        Dim msgRetWS As String
        Dim msgResultado As String
        Dim siglaWS As String
        Dim certificado As String
        '
        '  As variáveis do proxy devem ser informadas se necessário
        '
        '  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
        '
        Dim proxy As String
        Dim usuario As String
        Dim senha As String
        Dim licenca As String
        '
        Dim ambiente As Integer
        '
        ' define as variáveis que passam/recebem informações importantes
        '
        Dim ChaveNFe As String          ' chave da NF-e objeto de cancelamento
        Dim ProtAutNFe As String        ' protocolo de autorização de uso
        '
        '  parâmetros novos
        '
        Dim procCancNFe As String       ' estrturura XML que contém o pedido de cancelamento e a homologação do cancelamento,
        ' que deve ser mantido pelo emissor e distribuído ao destinatário.
        Dim nProtocoloCanc As String    ' número do protocolo de homomologação de cancelamento devolvido pela SEFA
        Dim dProtocoloCanc As String    ' data e hora de homologação do cancelamento
        Dim versao As String            'utilizado para escolha da versão do WS
        Dim Justificativa_cancelamento As String
        Dim Mensagem_entrada As String
        Dim Mensagem_saida As String
        Dim dhevento As String

        '
        '  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
        '

        proxy = ""
        usuario = ""
        senha = ""
        licenca = ""
        msgDados = ""
        msgRetWS = ""
        siglaWS = ""
        msgResultado = ""
        procCancNFe = ""
        nProtocoloCanc = ""
        dProtocoloCanc = ""
        certificado = ""
        versao = ""
        ChaveNFe = ""
        ProtAutNFe = ""
        dhevento = ""
        Justificativa_cancelamento = ""

        conn.Open()
        cmdNf.CommandText = "Select nf.*,fl.* from notas_fiscais nf (nolock), filiais fl (nolock) where nf.nota_fiscal = " & lnota_fiscal & " and nf.filial = fl.filial"
        dr = cmdNf.ExecuteReader()
        If dr.Read() Then
            certificado = dr.Item("Certificado_digital").ToString
            versao = Trim(dr.Item("Versao_xml").ToString)
            ChaveNFe = dr.Item("Chave_acesso").ToString
            Justificativa_cancelamento = dr.Item("Justificativa_cancelamento").ToString
            ambiente = dr.Item("Tipo_ambiente").ToString
            siglaWS = Função_Seleciona_siglaWS(dr.Item("dados_nfe_uf").ToString, dr.Item("Dados_nfe_forma_emissao").ToString + 1, dr.Item("Versao_xml").ToString)
            If BDretorna_campo("Parametros", "valor_texto", "parametro = 56") = "S" Then 'Horário de Verão
                dhevento = Format(dr.Item("status_ultimo_evento"), "yyyy-MM-ddTHH:mm:ss-02:00")
            Else
                dhevento = Format(dr.Item("status_ultimo_evento"), "yyyy-MM-ddTHH:mm:ss-03:00")
            End If

            ProtAutNFe = dr.Item("Protocolo_autorizacao").ToString
            licenca = dr.Item("Cte_chave_flexdocs").ToString
        End If
        dr.Close()
        conn.Close()
        cmdNf.Dispose()
        conn.Dispose()

        Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.canc.cancelaNF2G.html

        '
        ' referenciando a DLL em late binding
        ' não é necessário fazer o reference da DLL
        ' o intelisense não funciona
        '
        Dim objCTeUtil As Object

        objCTeUtil = CreateObject("CTe_Util.Util")

        '  trecho para instanciar a DLL em early binding
        '  necessario fazer o referece da DLL
        '
        'procCancNFe = objNFeUtil.CancelaNFEvento(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, ChaveNFe, ProtAutNFe, Justificativa_cancelamento, dhevento, nProtocoloCanc, dProtocoloCanc, proxy, usuario, senha, licenca)
        procCancNFe = objCTeUtil.CancelaCTEvento(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, ChaveNFe, ProtAutNFe, Justificativa_cancelamento, dhevento, nProtocoloCanc, dProtocoloCanc, proxy, usuario, senha, licenca)
        '
        '
        ' mostra mensagem XML enviada e a mensagem de retorno do WS
        '
        Mensagem_entrada = msgDados          ' string com a mensagem XML enviado ao WS

        Mensagem_saida = msgRetWS          ' string com a mensagem XML da resposta do WS

        If cStat = 135 Or cStat = 155 Then
            'MsgBox(msgResultado & Chr(13) & Chr(13) + "Protocolo de homologação de cancelamento: " + nProtocoloCanc + Chr(13) & Chr(13) + "Data e hora de homologação de cancelamento: " + dProtocoloCanc + Chr(13) & Chr(13) + "Grave o procCancNFe : " + procCancNFe, vbInformation, "Atenção: Cancelamento da NF-e")
            'Dim fTran As DbTransaction = conn.BeginTransaction

            'Inicia a transação com begin transaction
            'cmdNf.Transaction = fTran

            Dim sqlBuilder1 As New System.Text.StringBuilder
            'sqlBuilder1.Append("update romaneios set nota_fiscal = null where nota_fiscal = " & lnota_fiscal)
            'BDexecuta_query(sqlBuilder1)

            'Limpa o romaneio
            'cmdNf.CommandText = "update romaneios set nota_fiscal = null where nota_fiscal = " & lnota_fiscal
            'cmdNf.ExecuteNonQuery()

            'Limpa a negociacao
            Dim sqlBuilder2 As New System.Text.StringBuilder
            'sqlBuilder2.Append("update negociacoes set nota_fiscal = null where nota_fiscal = " & lnota_fiscal)
            'BDexecuta_query(sqlBuilder2)

            'cmdNf.CommandText = "update negociacoes set nota_fiscal = null where nota_fiscal = " & lnota_fiscal
            'cmdNf.ExecuteNonQuery()

            'Deixa os títulos como planejado
            'cmdNf.CommandText = "update Contas_a_receber set Situacao = 'P' where venda = " & lvenda)
            'cmdNf.ExecuteNonQuery()

            'Cancela Venda
            'cmdNf.CommandText = "update vendas set Cancelado = 1 where nota_fiscal = " & lnota_fiscal
            'cmdNf.ExecuteNonQuery()
            Dim sqlBuilder3 As New System.Text.StringBuilder
            'sqlBuilder3.Append("update vendas set Cancelado = 1 where nota_fiscal = " & lnota_fiscal)
            'BDexecuta_query(sqlBuilder3)

            'Limpa a venda
            'cmdNf.CommandText = "update vendas set nota_fiscal = null, Fechado = 0 where nota_fiscal = " & lnota_fiscal
            'cmdNf.ExecuteNonQuery()
            Dim sqlBuilder4 As New System.Text.StringBuilder
            'sqlBuilder4.Append("update vendas set nota_fiscal = null, Fechado = 0 where nota_fiscal = " & lnota_fiscal)
            'BDexecuta_query(sqlBuilder4)

            Dim sqlBuilder As New System.Text.StringBuilder
            sqlBuilder.Append("update notas_fiscais set Xml_cancelamento = '" & msgDados & "',Protocolo_cancelamento = '" & nProtocoloCanc & "',Data_hora_cancelamento='" & dProtocoloCanc & "' where nota_fiscal = " & lnota_fiscal)
            BDexecuta_query(sqlBuilder)

            'Cancela Nota Fiscal
            Call BDInclui_evento_nf("Cte Cancelado no Sefaz", lnota_fiscal, "CA")
        Else
            'Erro Cancelamento
            Call BDInclui_evento_nf(msgResultado, lnota_fiscal, "EC")
            'Seta Finalizado
            Call BDInclui_evento_nf("Cte Finalizado", lnota_fiscal, "FN")
        End If

    End Sub

    Private Sub NfeUtil_Inutiliza_nfe(ByVal lnota_fiscal As Long)
        Dim conn As DbConnection = Me.dbfactory.CreateConnection
        conn.ConnectionString = My.Settings.ConnectionString
        Dim cmdNf As DbCommand = conn.CreateCommand
        Dim dr As DbDataReader

        'Variáveis da função que inutiliza
        Dim lcertificado As String = ""
        Dim lsiglaWS As String = ""
        Dim lversao As String = ""
        Dim lCNPJ As String = ""
        Dim lano As Integer = 0
        Dim lserie As String = ""
        Dim lnrinicial As Long = 0
        Dim lnrfinal As Long = 0
        Dim lambiente As Integer = 0
        Dim lJustificativa As String = "Nf Inativada devido a acerto/falha operacional"
        '
        '  Inutiliza Número de NF-e
        '
        '  A funcionalidade deve ser utilizada para inutilizar um
        '  número de NF-e que não vai ser utilizada (atribuída) a
        '  NF-e, por salto de numeração, rejeição de NF-e, etc.
        '
        '  veja os detalhes da chamada em: http://www.flexdocs.com.br/guiaNFe/WS.canc.inutilizaNro2G.html
        '
        Dim msgDados As String
        Dim msgRetWS As String
        Dim msgResultado As String
        '
        '  As variáveis do proxy devem ser informadas se necessário
        '
        '  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
        '
        Dim proxy As String
        Dim usuario As String
        Dim senha As String
        Dim licenca As String
        '
        ' define as variáveis que passam/recebem informações importantes
        '
        Dim cUF As String = ""               ' código da UF do solicitante - Tabela IBGE
        Dim siglaUF As String = ""
        Dim modelo As String            ' modelo da NF-e (sempre 55)
        ' Observações
        ' só é permitida a inutilização de até 1000 números por vez
        ' se a inutilização for de um único número nInicial e nFinal devem
        Dim procInutNFe As String       ' estrturura XML que contém o pedido de inutilização e a homologação da inutilização,
        ' que deve ser mantido pelo emissor.
        Dim nProtocoloInut As String    ' número do protocolo de homomologação de Inutilização de numeraçãp devolvido pela SEFA
        Dim dProtocoloInut As String    ' data e hora de homologação da Inutilização de numeraçãp
        '
        '
        '  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
        '
        '
        proxy = ""
        usuario = ""
        senha = ""
        licenca = ""
        msgDados = ""
        msgRetWS = ""
        msgResultado = ""

        procInutNFe = ""
        nProtocoloInut = ""
        dProtocoloInut = ""

        ' informar com o assunto da certificado digital
        ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

        conn.Open()
        cmdNf.CommandText = "Select nf.Dados_nfe_uf,nf.dados_nfe_uf,nf.dados_nfe_uf_ibge,nf.tipo_ambiente,nf.dados_nfe_data_emissao,fl.Certificado_digital,nf.emitente_cnpj,nf.dados_nfe_serie,nf.dados_nfe_numero,fl.sefaz_virtual,fl.Cte_chave_flexdocs,nf.versao_xml from filiais fl (nolock),notas_fiscais nf (nolock) where nf.nota_fiscal = " & lnota_fiscal & " and nf.filial = fl.filial"
        dr = cmdNf.ExecuteReader()
        If dr.Read() Then
            lcertificado = dr.Item("Certificado_digital").ToString
            lCNPJ = Microsoft.VisualBasic.Left(dr.Item("emitente_cnpj").ToString, 2) + Microsoft.VisualBasic.Mid(dr.Item("emitente_cnpj").ToString, 4, 3) + Microsoft.VisualBasic.Mid(dr.Item("emitente_cnpj").ToString, 8, 3) + Microsoft.VisualBasic.Mid(dr.Item("emitente_cnpj").ToString, 12, 4) + Microsoft.VisualBasic.Right(dr.Item("emitente_cnpj").ToString, 2) ' CNPJ do emitente sem máscara de formatação
            lserie = Trim(dr.Item("dados_nfe_serie").ToString)
            lnrinicial = dr.Item("dados_nfe_numero").ToString
            lnrfinal = lnrinicial
            lano = Year(dr.Item("dados_nfe_data_emissao").ToString) - 2000
            lversao = Trim(dr.Item("versao_xml").ToString)
            lsiglaWS = Função_Seleciona_siglaWS(dr.Item("dados_nfe_uf").ToString, 1, lversao)
            cUF = Trim(dr.Item("dados_nfe_uf_ibge").ToString)
            licenca = dr.Item("Cte_chave_flexdocs").ToString
            lambiente = dr.Item("tipo_ambiente").ToString
            siglaUF = dr.Item("Dados_nfe_uf").ToString
        End If
        dr.Close()
        conn.Close()
        cmdNf.Dispose()
        conn.Dispose()

        ' o modelo da NF-e é sempre fixo em 57
        modelo = "57"
        '
        Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.canc.inutilizaNro2G.html
        '
        ' referenciando a DLL em late binding
        ' não é necessário fazer o reference da DLL
        ' o intelisense não funciona
        '
        Dim objCTeUtil As Object

        objCTeUtil = CreateObject("CTe_Util.Util")

        '
        '  trecho para instanciar a DLL em early binding
        '  necessario fazer o referece da DLL
        '
        'Dim objNFeUtil As NFe_Util_2G.Util
        '
        'Set objNFeUtil = New NFe_Util_2G.Util
        '
        'procInutNFe = objNFeUtil.InutilizaNroNF2G(lsiglaWS, lambiente, lcertificado, lversao, msgDados, msgRetWS, cStat, msgResultado, cUF, lano, lCNPJ, modelo, lserie, lnrinicial, lnrfinal, lJustificativa, nProtocoloInut, dProtocoloInut, proxy, usuario, senha, licenca)
        cStat = objCTeUtil.InutilizaNroCT(siglaUF, siglaUF, lambiente, lcertificado, lversao, msgDados, msgRetWS, msgResultado, procInutNFe, cUF, lano, lCNPJ, modelo, lserie, lnrinicial, lnrfinal, lJustificativa, proxy, usuario, senha, licenca)
        '
        ' mostra mensagem XML enviada e a mensagem de retorno do WS
        '
        'txtEntrada.Text = msgDados          ' string com a mensagem XML enviado ao WS

        'txtRetorno.Text = msgRetWS          ' string com a mensagem XML da resposta do WS

        If cStat = 102 Then
            Call BDInclui_evento_nf("Nfe Inutilizada", lnota_fiscal, "IU")
            Dim sqlBuilder As New System.Text.StringBuilder
            sqlBuilder.Append("update notas_fiscais set Protocolo_inutilizacao = '" & nProtocoloInut & "',Data_hora_inutilizacao='" & dProtocoloInut & "' where nota_fiscal = " & lnota_fiscal)
            BDexecuta_query(sqlBuilder)
        Else
            Call BDInclui_evento_nf("Erro: " & msgResultado, lnota_fiscal, "EI")
        End If
    End Sub

    Private Sub NfeUtil_carta_correcao(ByVal lnota_fiscal As Long, ByVal lnumero As Long)
        Dim conn As DbConnection = Me.dbfactory.CreateConnection
        conn.ConnectionString = My.Settings.ConnectionString
        Dim cmdNf As DbCommand = conn.CreateCommand
        Dim dr As DbDataReader

        Dim objCTeUtil As Object
        objCTeUtil = CreateObject("CTe_Util.Util")
        '
        '  Carta de Correção eletrônica
        '
        '  Exemplo de uso da funcionalidade de carta de correção eletrônica
        '
        '  veja detalhes da funcionalidade em: http://www.flexdocs.com.br/guiaNFe/WS.evento.CCe.html
        '
        Dim msgDados As String
        Dim msgRetWS As String
        Dim msgResultado As String
        Dim siglaWS As String
        Dim certificado As String
        '
        '  As variáveis do proxy devem ser informadas se necessário
        '
        '  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
        '
        Dim proxy As String
        Dim usuario As String
        Dim senha As String
        Dim licenca As String
        '
        Dim ambiente As Integer
        '
        ' define as variáveis que passam informações para a DLL
        '
        Dim versao As String            ' utilizado para escolha da versão do WS, informar "3.00"
        Dim ChaveCte As String          ' chave da CT-e objeto de carta de correção eletrônica
        Dim XmlCorrecao As String      ' texto da correção - string com até 1000 caracteres
        Dim dhCorrecao As String        ' data e hora da correção
        Dim nCCe As Long                ' número da carta de correção, deve ser um número sequencial iniciado em 1, o valor máximo é 20
        Dim descEventoAcentuado As Long ' indicardor de acentuação da descrição do evento e das condições de uso, deve ser informado com 0-não/1-sim
        ' indicar com 0 para as UF que não aceitam acento como é o caso do MT
        ' IMPORTANTE: o controle da acentuação do texto da correção é da aplicação do usuário, este indicador serve
        ' apenas para que a DLL informe os campos descEvento e xCondUso sem acentuaçã.
        '
        '  parâmetros que devolvem informações
        '
        Dim procCCe As String           ' estrturura XML que contém a carta de correção eletrônica e registro do evento da carta de correção eletrônica,
        ' que deve ser mantido pelo emissor e distribuído ao destinatário.
        Dim nProtocoloCCe As String    ' número do protocolo de  registro do evento da carta de correção eletrônica devolvido pela SEFA
        Dim dProtocoloCCe As String    ' data e hora de  registro do evento da carta de correção eletrônica

        Dim lNota_fiscal_carta_correcao As Long
        Dim emi_CNPJ As String = "" ' CNPJ do emitente sem máscara de formatação
        '
        '
        '  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
        '
        '
        proxy = ""
        usuario = ""
        senha = ""
        licenca = ""
        msgDados = ""
        msgRetWS = ""
        msgResultado = ""
        procCCe = ""
        certificado = ""
        nProtocoloCCe = ""
        dProtocoloCCe = ""
        versao = ""
        siglaWS = ""
        ChaveCte = ""
        XmlCorrecao = ""
        dhCorrecao = ""

        conn.Open()
        cmdNf.CommandText = "Select nf.Emitente_cnpj,cce.Nota_fiscal_carta_correcao,nf.Versao_xml,nf.dados_nfe_uf,nf.Dados_nfe_forma_emissao,fl.certificado_digital,nf.Chave_acesso,nf.Tipo_ambiente,cce.data_inclusao,cce.numero,cce.versao,cce.texto,fl.sefaz_virtual,fl.Cte_chave_flexdocs from notas_fiscais nf (nolock), filiais fl (nolock),notas_fiscais_cartas_correcao cce (nolock) where cce.numero = " & lnumero & " and nf.nota_fiscal = cce.nota_fiscal and nf.nota_fiscal = " & lnota_fiscal & " and nf.filial = fl.filial"
        dr = cmdNf.ExecuteReader()
        If dr.Read() Then
            certificado = dr.Item("Certificado_digital").ToString
            ChaveCte = dr.Item("Chave_acesso").ToString
            ambiente = dr.Item("Tipo_ambiente").ToString
            'siglaWS = dr.Item("sefaz_virtual").ToString
            siglaWS = Função_Seleciona_siglaWS(dr.Item("dados_nfe_uf").ToString, dr.Item("Dados_nfe_forma_emissao").ToString + 1, dr.Item("Versao_xml").ToString)
            If BDretorna_campo("Parametros", "valor_texto", "parametro = 56") = "S" Then 'Horário de Verão
                dhCorrecao = Format(dr.Item("data_inclusao"), "yyyy-MM-ddTHH:mm:ss-02:00") ' data de emissão
            Else
                dhCorrecao = Format(dr.Item("data_inclusao"), "yyyy-MM-ddTHH:mm:ss-03:00") ' data de emissão
            End If
            nCCe = dr.Item("Numero").ToString
            versao = Trim(dr.Item("Versao_xml").ToString)
            XmlCorrecao = dr.Item("Texto").ToString
            licenca = dr.Item("Cte_chave_flexdocs").ToString
            lNota_fiscal_carta_correcao = dr.Item("Nota_fiscal_carta_correcao").ToString
            emi_CNPJ = Microsoft.VisualBasic.Left(dr.Item("Emitente_cnpj").ToString, 2) + Microsoft.VisualBasic.Mid(dr.Item("Emitente_cnpj").ToString, 4, 3) + Microsoft.VisualBasic.Mid(dr.Item("Emitente_cnpj").ToString, 8, 3) + Microsoft.VisualBasic.Mid(dr.Item("Emitente_cnpj").ToString, 12, 4) + Microsoft.VisualBasic.Right(dr.Item("Emitente_cnpj").ToString, 2) ' CNPJ do emitente sem máscara de formatação
        End If
        dr.Close()
        '
        'Pega a Carta de Correção
        '
        cmdNf.CommandText = "Select cccte.correcao,CteCampos.Grupo,CteCampos.Campo from Notas_fiscais_cartas_correcao_cte_campos cccte,Notas_fiscais_cc_cte_campos CteCampos where cccte.Nota_fiscal_cc_cte_campo = CteCampos.Nota_fiscal_cc_cte_campo and cccte.Nota_fiscal_carta_correcao = " & lNota_fiscal_carta_correcao
        dr = cmdNf.ExecuteReader()
        Do While dr.Read
            XmlCorrecao = XmlCorrecao & objCTeUtil.geraInfCorrecao(dr.Item("Grupo").ToString, dr.Item("Campo").ToString, dr.Item("Correcao").ToString, "")
        Loop
        dr.Close()
        conn.Close()
        cmdNf.Dispose()
        conn.Dispose()

        descEventoAcentuado = 1             ' indicador de acentuação da descrição do evento e das condições de uso, deve ser informado com 0-não/1-sim

        Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.evento.CCe.html

        procCCe = objCTeUtil.EnviaCCe(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, ChaveCte, XmlCorrecao, descEventoAcentuado, nCCe, dhCorrecao, nProtocoloCCe, dProtocoloCCe, proxy, usuario, senha, licenca)
        '
        ' mostra mensagem XML enviada e a mensagem de retorno do WS
        '
        'MsgBox(msgDados)          ' string com a mensagem XML enviado ao WS

        'MsgBox(msgRetWS)          ' string com a mensagem XML da resposta do WS

        If cStat = 135 Then
            'Atualiza o status da nota fiscal
            Dim sqlBuilder As New System.Text.StringBuilder
            sqlBuilder.Append("update notas_fiscais_cartas_correcao set Protocolo = '" & nProtocoloCCe & "',Data_hora = '" & dProtocoloCCe & "' where nota_fiscal = " & lnota_fiscal & " and numero = " & lnumero)
            BDexecuta_query(sqlBuilder)
            'Salva o XML
            Dim larquivo_xml As String = My.Application.Info.DirectoryPath & "\Cte\" & emi_CNPJ & "\" & ChaveCte & "-" & nCCe & ".xml"
            If IO.File.Exists(larquivo_xml) Then System.IO.File.Delete(larquivo_xml)
            Dim objStream1 As New System.IO.FileStream(larquivo_xml, IO.FileMode.Create)
            Dim Arq1 As New System.IO.StreamWriter(objStream1)
            Arq1.Write(procCCe)
            Arq1.Close()
            '
            Call BDInclui_evento_nf("Correção da carta " & nCCe & " executada com sucesso", lnota_fiscal, "CO")
        Else
            Call BDInclui_evento_nf("Correção da carta " & nCCe & " falhou:" & msgDados & " " & msgResultado, lnota_fiscal, "EO")
        End If
        'Call BDInclui_evento_nf("Cte Finalizado", lnota_fiscal, "FN")
    End Sub

    Private Sub Timer_nfes_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer_ctes.Elapsed
        Timer_ctes.Enabled = False
        EventLog1.WriteEntry("Verificando os ctes " & DateTime.Now)
        'arquivoWS.WriteLine("Verificando os nfs " & DateTime.Now)
        Call Processa_ctes()
        Timer_ctes.Enabled = True
    End Sub

    Public Sub Cte_gera_pdf(ByVal lnota_fiscal As Long, ByVal lemi_CNPJ As String, ByVal lChaveNFe As String, ByVal lapelido_filial As String)
        '
        ' declaração das variáveis que serão utilizadas na passagem de parâmetros da DLL
        '
        Dim XML As String                 ' informar o XML do CT-e da versão 1.04
        Dim logo As String = ""                ' nome arquivo do logotipo
        Dim quadroRecibo As String        ' posicão de impressão do quadroRecibo [S]uperior [I]nferior
        Dim visualizar As String          ' data e número do registro do DPEC
        Dim parametros As String          ' parametros da geracao do PDF
        Dim cResultado As Long            ' código deretorno da chamada da DLL
        Dim msgResultado As String        ' literal com resultado da chamada da DLL
        Dim OrigDadosEmissor As String = ""

        On Error GoTo 0
        '   Carrega o conteúdo do nome do arquivo em XMLString
        '
        'Abre o arquivo XML para consultar no sefaz
        '
        Dim nomeArquivoCTe As String
        nomeArquivoCTe = My.Application.Info.DirectoryPath & "\cte\" & lemi_CNPJ & "\" & lChaveNFe & ".xml"

        Dim objStream As New System.IO.FileStream(nomeArquivoCTe, IO.FileMode.Open)
        Dim Arq As New System.IO.StreamReader(objStream)
        XML = Arq.ReadLine
        Arq.Close()
        '
        'Verifica se tem e seta a imagem para imprimir na danfe 
        '
        Dim nomeArquivoBmp As String
        nomeArquivoBmp = My.Application.Info.DirectoryPath & "\" & lapelido_filial & ".jpg"
        If IO.File.Exists(nomeArquivoBmp) Then OrigDadosEmissor = nomeArquivoBmp
        'OrigDadosEmissor = ""           ' origem dos dados do emissor no XML, possibilidades:
        '  sem conteúdo -> os dados do emissor serão obtidos do XML;
        '  nome arquivo -> a imagem informada irá ocupar todo o quadro dados do emitente;
        '  literal [SEM DADOS EMITENTE] -> nenhum dado será impresso no quado dados do emitente;
        quadroRecibo = "S"              ' quadro do recibo no topo "S" ou rodape "I"
        visualizar = "N"                ' visualizar PDF "S" ou "N"

        ' Parâmetro, valores válidos:
        ' [RODAPE=texto do rodape] -> imprime o "texto do rodape" informado no RODAPE;
        ' [PASTA=] -> indica a pasta de gravação do PDF;
        ' [VISUALIZAR] -> indica visualização da PDF;
        ' [ARQUIVO=nomeArquivo] -> grava o PDF com o nome indicado;
        ' [MENSAGEM=texto da mensagem] -> imprime o "texto da mensagem" informado no corpo do DANFE;
        ' [IMPRIMIR=n] -> imprime n cópias do DACTE
        '
        parametros = "[RODAPE=Emitido por SisAgro - Sistemas para o Agronegócio. www.sisprime.com.br. (44) 98811-6666.][ARQUIVO=" & lChaveNFe & ".PDF][PASTA=" & My.Application.Info.DirectoryPath & "\cte\" & lemi_CNPJ & "\][[VISUALIZAR=N]"
        'parametros = "[RODAPE=" & OrigDadosEmissor & ".][ARQUIVO=" & lChaveNFe & ".PDF][PASTA=" & My.Application.Info.DirectoryPath & "\cte\" & lemi_CNPJ & "\][[VISUALIZAR=N]"

        cResultado = 0
        msgResultado = ""
        '
        ' instancia a DLL - late binding
        '
        Dim objCTeUtil As Object
        '
        objCTeUtil = CreateObject("CTe_Util.Util")
        '
        ' chama DLL
        cResultado = objCTeUtil.geraPdfDACTE(XML, OrigDadosEmissor, quadroRecibo, visualizar, parametros, msgResultado)
        '
        '  tratar retorno
        '
        BDInclui_evento_nf(msgResultado, lnota_fiscal, "GP")
        '
        '  liberar DLL
        '
        objCTeUtil = Nothing
    End Sub

End Class
