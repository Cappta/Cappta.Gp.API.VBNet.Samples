Imports System.Configuration
Imports System.Text
Imports System.Threading
Imports Cappta.Gp.Api.Com
Imports Cappta.Gp.Api.Com.Model
Imports Cappta.Gp.Api.Com.Sample.VB.My.Resources

Public Class FormularioSampleVB
    Dim iteracaoTef As IIteracaoTef

    Private Const TIPO_VIA_TODAS As Integer = 1

    Private Const TIPO_VIA_CLIENTE As Integer = 2

    Private Const TIPO_VIA_LOJA As Integer = 3

    Dim tipoVia As Integer = TIPO_VIA_TODAS

    Dim processandoPagamento As Double

    Dim detailsCrediario = New DetalhesCrediario

    Dim valor As Double

    Dim tiposParcelamento As Dictionary(Of Integer, TipoParcelamento)

    Dim quantidadeCartoes As Integer = 0

    Dim sessaoMultiTefEmAndamento As Boolean

    Dim INTERVALO_MILISEGUNDOS As Integer = 500

    Dim cliente As New ClienteCappta

    Dim resultadoDaAutenticacao As Integer

    Dim resultado As Integer

    Private Sub FormularioSampleVB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AutenticarPdv()
        ConfigurarModoIntegracao(True)
        IniciarControles()
    End Sub

    Private Sub IniciarControles()

        Dim tiposParcelamento = New Dictionary(Of Integer, TipoParcelamento)()
        tiposParcelamento.Add(0, TipoParcelamento.Administrativo)
        tiposParcelamento.Add(1, TipoParcelamento.Loja)

        ComboBoxTipoParcelamentoPagamentoCredito.SelectedIndex = 0
        ComboBoxTipoInformacaoPinpad.DataSource = [Enum].GetValues(GetType(TipoInformacaoPinpad))

    End Sub

    Private Sub AutenticarPdv()


        Dim chaveAutenticacao = ConfigurationManager.AppSettings("ChaveAutenticacao")

        If String.IsNullOrWhiteSpace(chaveAutenticacao) Then
            InvalidarAutenticacao("Chave de autenticação inválida")
        End If

        Dim cnpj = ConfigurationManager.AppSettings("Cnpj")
        If String.IsNullOrWhiteSpace(cnpj) Or cnpj.Length <> 14 Then
            InvalidarAutenticacao("CNPJ inválido")
        End If

        Dim pdv = ConfigurationManager.AppSettings("Pdv")
        If Int32.TryParse("pdv", pdv = False Or pdv = 0) Then
            InvalidarAutenticacao("PDV Inválido")
        End If
        resultadoDaAutenticacao = cliente.AutenticarPdv(cnpj, pdv, chaveAutenticacao)
        resultadoDaAutenticacao = cliente.AutenticarPdv(cnpj, pdv, chaveAutenticacao)
        If resultadoDaAutenticacao = 0 Then
            Return
        End If

        Dim mensagem As String = Mensagens.ResourceManager.GetString(String.Format("RESULTADO_CAPPTA_{0}", resultadoDaAutenticacao))
        ExibirMensagemAutenticacaoInvalida(resultadoDaAutenticacao)
    End Sub

    Private Sub ExibirMensagemAutenticacaoInvalida(resultadoAutenticacao As Integer)
        Dim mensagem = Mensagens.ResourceManager.GetString(String.Format("RESULTADO_CAPPTA_{0}", resultadoAutenticacao))
        MessageBox.Show(mensagem, "SAMPLE API COM", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

    End Sub

    Private Sub InvalidarAutenticacao(mensagemErro As String)
        CriarMensagemErroJanela(String.Format("RESULTADO_CAPPTA_{0}", resultadoDaAutenticacao))
        Environment.Exit(0)
    End Sub

    Private Sub ConfigurarModoIntegracao(exibirInterface As Boolean)

        Dim configs As IConfiguracoes = New Configuracoes
        configs.ExibirInterface = exibirInterface

        resultado = cliente.Configurar(configs)
        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
            Return
        End If
    End Sub

    ' Métodos de Pagamento
    Private Sub ExecutarDebito_Click(sender As Object, e As EventArgs) Handles ExecutarDebito.Click

        If DeveIniciarMultiCartoes() Then
            IniciarMulticartoes()
        End If

        Dim valor As Decimal = NumericUpDownValorPagamentoDebito.Value

        Dim resultado As Int32
        resultado = cliente.PagamentoDebito(valor)

        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
            Return
        End If

        processandoPagamento = True
        IterarOperacaoTef()
    End Sub

    Private Sub ExecutarCredito_Click_1(sender As Object, e As EventArgs) Handles ExecutarCredito.Click
        If DeveIniciarMultiCartoes() Then
            IniciarMulticartoes()
        End If

        valor = NumericUpDownValorPagamentoCredito.Value

        Dim details As IDetalhesCredito = New DetalhesCredito
        details.QuantidadeParcelas = NumericUpDownQuantidadeParcelasPagamentoCredito.Value
        details.TipoParcelamento = ComboBoxTipoParcelamentoPagamentoCredito.SelectedIndex
        details.TransacaoParcelada = RadioButtonPagamentoCreditoComParcelas.Checked


        resultado = cliente.PagamentoCredito(valor, details)
        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
        End If
        processandoPagamento = True
        IterarOperacaoTef()
    End Sub

    Private Sub ExecutarCrediario_Click(sender As Object, e As EventArgs) Handles ExecutarCrediario.Click
        Dim valor As Double = NumericUpDownQuantidadeParcelasPagamentoCrediario.Value

        detailsCrediario.QuantidadeParcelas = NumericUpDownQuantidadeParcelasPagamentoCrediario.Value

        If DeveIniciarMultiCartoes() Then
            IniciarMulticartoes()
        End If

        cliente.PagamentoCrediario(valor, detailsCrediario)
        Dim resultado As Integer = cliente.PagamentoCrediario(valor, detailsCrediario)
        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)

        End If

        processandoPagamento = True
        IterarOperacaoTef()

    End Sub

    Private Sub ExecutarTicketCar_Click(sender As Object, e As EventArgs) Handles ExecutarTicketCar.Click
        resultado = 0

        Dim valor As Double = NumericUpDownValorPagamentoTicketCar.Value
        Dim detalhesTicketCar As New DetalhesPagamentoTicketCarPessoaFisica

        detalhesTicketCar.NumeroReciboFiscal = TextBoxDocumentoFiscal.Text
        detalhesTicketCar.NumeroSerialECF = TextBoxNumeroSerial.Text

        resultado = cliente.PagamentoTicketCarPessoaFisica(valor, detalhesTicketCar)
        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
            Return
        End If

        processandoPagamento = True
        IterarOperacaoTef()

    End Sub

    Function DeveIniciarMultiCartoes()
        Return sessaoMultiTefEmAndamento = False And RadioButtonUsarMultiTef.Checked
    End Function

    Private Sub IniciarMulticartoes()
        quantidadeCartoes = NumericUpDownQuantidadeDePagamentosMultiTef.Value
        sessaoMultiTefEmAndamento = True
        cliente.IniciarMultiCartoes(quantidadeCartoes)
    End Sub

    'Region Métodos Administrativos
    Private Sub ExecutarReimpressao_Click(sender As Object, e As EventArgs) Handles ExecutarReimpressao.Click

        If sessaoMultiTefEmAndamento = True Then
            CriarMensagemErroJanela("Não é possível reimprimir um cupom com uma sessão multitef em andamento.")
        End If

        Dim resultado As Integer = RadioButtonReimprimirUltimoCupom.Checked

        cliente.ReimprimirUltimoCupom(tipoVia)
        cliente.ReimprimirCupom(NumericUpDownNumeroControleReimpressaoCupom.Value.ToString("00000000000"), tipoVia)

        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
        End If

        processandoPagamento = False
        IterarOperacaoTef()
    End Sub

    Private Sub ExecutarCancelamento_Click(sender As Object, e As EventArgs) Handles ExecutarCancelamento.Click
        If sessaoMultiTefEmAndamento = True Then
            CriarMensagemErroJanela("Não é possível cancelar um pagamento com uma sessão multitef em andamento.")
            Return
        End If

        Dim senhaAdministrativa = TextBoxSenhaAdministrativaCancelamento.Text

        If String.IsNullOrEmpty(senhaAdministrativa) Then
            CriarMensagemErroJanela("A senha administrativa não pode ser vazia.")
            Return
        End If

        Dim numeroControle = NumericUpDownNumeroControleCancelamento.Value.ToString("00000000000")

        Dim resultado = cliente.CancelarPagamento(senhaAdministrativa, numeroControle)
        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
            Return
        End If

        processandoPagamento = False
        IterarOperacaoTef()

    End Sub

    Private Sub CriarMensagemErroJanela(mensagem As String)

        MessageBox.Show("Erro", mensagem)

    End Sub

    Private Sub CriarMensagemErroPainel(resultado As Integer)
        Dim mensagem As String = Mensagens.ResourceManager.GetString(String.Format("RESULTADO_CAPPTA_{0}", resultado))
        If String.IsNullOrEmpty(mensagem) Then
            mensagem = "Não foi possível executar a operação."
            AtualizarResultado(String.Format("{0}. Código de erro {1}", mensagem, resultado))
        End If

        AtualizarResultado(mensagem)
        TextBoxResultado.Text = mensagem
        TextBoxResultado.Update()
    End Sub

    'Region Método IterarOperacaoTef
    Public Sub IterarOperacaoTef()

        If RadioButtonUsarMultiTef.Enabled Then
            DesabilitarControlesMultiTef()
        End If

        Dim iteracaoTef As IIteracaoTef

        Do
            iteracaoTef = cliente.IterarOperacaoTef()

            If TypeOf iteracaoTef Is IMensagem Then
                ExibirMensagem(iteracaoTef)
                Thread.Sleep(INTERVALO_MILISEGUNDOS)
            End If

            If TypeOf iteracaoTef Is IRequisicaoParametro Then
                RequisitarParametros(iteracaoTef)
            End If

            If TypeOf iteracaoTef Is IRespostaTransacaoPendente Then
                ResolverTransacaoPendente(iteracaoTef)
            End If

            If TypeOf iteracaoTef Is IRespostaOperacaoRecusada Then
                ExibirDadosOperacaoRecusada(iteracaoTef)
            End If

            If TypeOf iteracaoTef Is IRespostaOperacaoAprovada Then
                ExibirDadosOperacaoAprovada(iteracaoTef)
                FinalizarPagamento()
            End If

        Loop While OperacaoNaoFinalizada(iteracaoTef)

        If sessaoMultiTefEmAndamento = False Then
            HabilitarControlesMultiTef()
        End If

    End Sub

    Private Sub HabilitarControlesMultiTef()

        HabilitarControle(RadioButtonUsarMultiTef)
        HabilitarControle(RadioButtonNaoUsarMultiTef)
        HabilitarControle(NumericUpDownQuantidadeDePagamentosMultiTef)

    End Sub

    Private Sub HabilitarBotoes()
        HabilitarControle(RadioButtonUsarMultiTef)
        HabilitarControle(ExecutarCredito)
        HabilitarControle(ExecutarCrediario)
        HabilitarControle(ExecutarReimpressao)
        HabilitarControle(ExecutarCancelamento)
    End Sub

    Private Sub HabilitarControle(controle As Control)
        controle.Enabled = True
        controle.Update()
    End Sub

    Private Sub FinalizarPagamento()
        If processandoPagamento = False Then
            Return
        End If

        If sessaoMultiTefEmAndamento Then

            If quantidadeCartoes > 0 Then
                Return
            End If
        End If
        Dim mensagem As String = GerarMensagemTransacaoAprovada()
        processandoPagamento = False
        sessaoMultiTefEmAndamento = False

        Dim result As DialogResult = MessageBox.Show(mensagem.ToString(), "Sample API COM", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        If result = DialogResult.OK Then
            cliente.ConfirmarPagamentos()

        Else
            cliente.DesfazerPagamentos()
        End If

    End Sub

    Private Function GerarMensagemTransacaoAprovada()

        Dim mensagem As String = "Clique em OK para confirmar a transação e em Cancelar para desfaze-la"

        Return mensagem
    End Function

    Private Sub ExibirDadosOperacaoAprovada(resposta As IRespostaOperacaoAprovada)

        Dim mensagemAprovada As New StringBuilder()

        If String.IsNullOrEmpty(resposta.CupomCliente) = False Then
            mensagemAprovada.AppendLine(resposta.CupomCliente.Replace("\", String.Empty)).AppendLine().AppendLine()
        End If

        If String.IsNullOrEmpty(resposta.CupomLojista) = False Then
            mensagemAprovada.Append(resposta.CupomLojista.Replace("\", String.Empty)).AppendLine()
        End If

        If String.IsNullOrEmpty(resposta.CupomReduzido) = False Then
            mensagemAprovada.Append(resposta.CupomReduzido.Replace("\", String.Empty)).AppendLine()
        End If

        AtualizarResultado(mensagemAprovada.ToString())

    End Sub

    Private Sub ExibirDadosOperacaoRecusada(resposta As IRespostaOperacaoRecusada)
        AtualizarResultado(String.Format("Código
             {0}{1}{2}", resposta.CodigoMotivo, Environment.NewLine, resposta.Motivo))
    End Sub

    Private Sub ResolverTransacaoPendente(transacaoPendente As IRespostaTransacaoPendente)
        Dim mensagemTransacoesPendentes As StringBuilder = New StringBuilder()
        mensagemTransacoesPendentes.AppendLine(transacaoPendente.Mensagem)
        For Each transacao In transacaoPendente.ListaTransacoesPendentes
            mensagemTransacoesPendentes.AppendLine($"Número de Controle
             {transacao.NumeroControle}")
            mensagemTransacoesPendentes.AppendLine($"Bandeira
             {transacao.NomeBandeiraCartao}")
            mensagemTransacoesPendentes.AppendLine($"Adquirente
             {transacao.NomeAdquirente}")
            mensagemTransacoesPendentes.AppendLine($"Valor
             {transacao.Valor}")
            mensagemTransacoesPendentes.AppendLine($"Data
             {transacao.DataHoraAutorizacao}")
        Next

        Dim input As String = Microsoft.VisualBasic.Interaction.InputBox(mensagemTransacoesPendentes.ToString())
        cliente.EnviarParametro(input, String.IsNullOrWhiteSpace(input))

    End Sub

    Private Sub RequisitarParametros(requisicaoParametros As IRequisicaoParametro)

        Dim input As String = InputBox(requisicaoParametros.Mensagem + Environment.NewLine + Environment.NewLine)
        Dim parametro As Integer
        If (String.IsNullOrWhiteSpace(input)) Then
            parametro = 2
        Else
            parametro = 1
        End If
        cliente.EnviarParametro(input, parametro)

    End Sub

    Private Sub ExibirMensagem(mensagem As Mensagem)
        AtualizarResultado(mensagem.Descricao)
    End Sub

    Function OperacaoNaoFinalizada(iteracaoTef As IIteracaoTef)

        Return iteracaoTef.TipoIteracao <> 1 And iteracaoTef.TipoIteracao <> 2
    End Function

    Private Sub AtualizarResultado(mensagem As String)
        TextBoxResultado.Text = mensagem
        TextBoxResultado.Update()
    End Sub

    Private Sub DesabilitarControlesMultiTef()
        DesabilitarControle(RadioButtonUsarMultiTef)
        DesabilitarControle(RadioButtonNaoUsarMultiTef)
        DesabilitarControle(NumericUpDownQuantidadeDePagamentosMultiTef)
    End Sub

    Private Sub DesabilitarControle(controle As Control)
        controle.Enabled = False
        controle.Update()
    End Sub

    Private Sub DesabilitarBotoes()
        DesabilitarControle(ExecutarDebito)
        DesabilitarControle(ExecutarCredito)
        DesabilitarControle(ExecutarCrediario)
        DesabilitarControle(ExecutarReimpressao)
        DesabilitarControle(ExecutarCancelamento)
    End Sub
    Private Sub ButtonSolicitarInformacaoPinpad_Click(sender As Object, e As EventArgs) Handles ButtonSolicitarInformacaoPinpad.Click
        Dim tipoDeEntrada = (ComboBoxTipoInformacaoPinpad.SelectedValue)

        Dim requisicaoPinpad As IRequisicaoInformacaoPinpad = New RequisicaoInformacaoPinpad()
        requisicaoPinpad.TipoInformacaoPinpad = tipoDeEntrada

        Dim informacaoPinpad = cliente.SolicitarInformacoesPinpad(requisicaoPinpad)
        AtualizarResultado(informacaoPinpad)
    End Sub

    Private Sub RadioButtonReimprimirViaCliente_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonReimprimirViaCliente.CheckedChanged
        tipoVia = TIPO_VIA_CLIENTE
    End Sub

    Private Sub RadioButtonReimprimirViaLoja_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonReimprimirViaLoja.CheckedChanged
        tipoVia = TIPO_VIA_LOJA
    End Sub

    Private Sub RadioButtonReimprimirTodasVias_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonReimprimirTodasVias.CheckedChanged
        tipoVia = TIPO_VIA_TODAS
    End Sub

    Private Sub RadioButtonNaoReimprimirUltimoCupom_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonNaoReimprimirUltimoCupom.CheckedChanged
        LabelNumeroControleReimpressaoCupom.Show()
        NumericUpDownNumeroControleReimpressaoCupom.Show()
    End Sub

    Private Sub RadioButtonReimprimirUltimoCupom_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonReimprimirUltimoCupom.CheckedChanged
        LabelNumeroControleReimpressaoCupom.Hide()
        NumericUpDownNumeroControleReimpressaoCupom.Hide()
    End Sub

    Private Sub RadioButtonPagamentoCreditoSemParcelas_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonPagamentoCreditoSemParcelas.CheckedChanged
        ComboBoxTipoParcelamentoPagamentoCredito.Hide()
        LabelTipoParcelamentoPagamentoCredito.Hide()
        NumericUpDownQuantidadeParcelasPagamentoCredito.Hide()
        LabelQuantidadeParcelasPagamentoCredito.Hide()
    End Sub

    Private Sub RadioButtonPagamentoCreditoComParcelas_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonPagamentoCreditoComParcelas.CheckedChanged
        ComboBoxTipoParcelamentoPagamentoCredito.Show()
        LabelTipoParcelamentoPagamentoCredito.Show()
        NumericUpDownQuantidadeParcelasPagamentoCredito.Show()
        LabelQuantidadeParcelasPagamentoCredito.Show()
    End Sub

    Private Sub NumericUpDownQuantidadeDePagamentosMultiTef_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownQuantidadeDePagamentosMultiTef.ValueChanged
        quantidadeCartoes = NumericUpDownQuantidadeDePagamentosMultiTef.Value
    End Sub

    Private Sub RadioButtonNaoUsarMultiTef_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonNaoUsarMultiTef.CheckedChanged
        If RadioButtonNaoUsarMultiTef.Checked = False Then
            Return
        End If

        sessaoMultiTefEmAndamento = False
        LabelQuantidadeDePagamentosMultiTef.Hide()
        NumericUpDownQuantidadeDePagamentosMultiTef.Hide()
        cliente.DesfazerPagamentos()
    End Sub

    Private Sub RadioButtonUsarMultiTef_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonUsarMultiTef.CheckedChanged
        If RadioButtonUsarMultiTef.Checked = False Then Return

        LabelQuantidadeDePagamentosMultiTef.Show()
        NumericUpDownQuantidadeDePagamentosMultiTef.Show()
    End Sub

    Private Sub RadioButtonInterfaceVisivel_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButtonInterfaceVisivel.CheckedChanged

        If RadioButtonInterfaceInvisivel.Checked = False Then
            Return
        End If
        ConfigurarModoIntegracao(False)

    End Sub

    Private Sub RadioButtonInterfaceInvisivel_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonInterfaceInvisivel.CheckedChanged

        If RadioButtonInterfaceVisivel.Checked = False Then
            Return
        End If

        ConfigurarModoIntegracao(True)
    End Sub
End Class




