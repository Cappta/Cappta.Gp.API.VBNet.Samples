<h1>Integração VB.Net</h1>

A Dll da Cappta foi desenvolvida utilizando as melhores práticas de programação e desenvolvimento de software. Utilizamos o padrão COM pensando justamente na integração entre aplicações construídas em várias linguagens. 

Obs: Durante a instalação do CapptaGpPlus o mesmo encarrega-se de registrar a DLL em seu computador.

<h3>Primeira etapa para integração.</h3></br>

Tempo estimado de 01:00 hora

 A primeira etapa consiste na importação do componente (dll) para dentro do projeto.</br>
 No Visual Studio, abra a Solution Explorer, vá em referências e adicione a DLL.
 
	
A primeira função a ser utilizada é **AutenticarPdv()**.</br>
     
Para autenticar é necessário os seguintes dados: CNPJ, PDV e chave de autenticação, estes dados são os mesmos fornecidos durante a instalação do GP.</br>
	
Chave: AAAAAAAAAAAAAAAA00000000000000A0 </br>

OBS: aqui utilizamos um xml para guardar os dados de autenticação.

Em App.config:
```javascript
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.0" />
  </startup>
  <appSettings>
    <add key="ChaveAutenticacao" value="AAAAAAAAAAAAAAAA00000000000000A0" />
    <add key="Cnpj" value="00000000000000" />
    <add key="Pdv" value="6" />
    <add key="ClientSettingsProvider.ServiceUri" value="" />
  </appSettings>
  <system.web>
    <membership defaultProvider="ClientAuthenticationMembershipProvider">
      <providers>
        <add name="ClientAuthenticationMembershipProvider" type="System.Web.ClientServices.Providers.ClientFormsAuthenticationMembershipProvider, System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" serviceUri="" />
      </providers>
    </membership>
    <roleManager defaultProvider="ClientRoleProvider" enabled="true">
      <providers>
        <add name="ClientRoleProvider" type="System.Web.ClientServices.Providers.ClientRoleProvider, System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" serviceUri="" cacheTimeout="86400" />
      </providers>
    </roleManager>
  </system.web>
</configuration>
```
Aconselhamos sempre deixar visivel e de fácil acesso para o usuário onde configurar o sistema para utilização do CapptaGpPlus.

```javascript

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
        If resultadoDaAutenticacao = 0 Then
            Return
        End If

        Dim mensagem As String = Mensagens.ResourceManager.GetString(String.Format("RESULTADO_CAPPTA_{0}", resultadoDaAutenticacao))
        ExibirMensagemAutenticacaoInvalida(resultadoDaAutenticacao)
    End Sub
```
O resultado para autenticação com sucesso é: 0

<h1>Primeiro esforço.</h1>
	Toda vez que realizar uma ação com o GP, vai perceber que ele começa a exibir o código 2 para autenticação, não se preocupe é assim mesmo, para recuperar os estados do GP, vamos direto para a etapa 3.

<h1> Etapa 2 </h1>

Tempo estimado de 00:40 minutos

Temos duas formas de integração, a visivel, onde a interação com o usuário fica por conta da Cappta, e a invisivel onde o form pode ser personalizado.


<h3>Para configurar o modo de integração</h3>

```javascript
 Private Sub ConfigurarModoIntegracao(exibirInterface As Boolean)

        Dim configs As IConfiguracoes = New Configuracoes
        configs.ExibirInterface = exibirInterface

        resultado = cliente.Configurar(configs)
        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
            Return
        End If
    End Sub
```

As mensagens de exceção, ficam em um Resources File:

RESULTADO_CAPPTA_1	Não autorizado. Por favor, realize a autenticação para utilizar o CapptaGpPlus

RESULTADO_CAPPTA_10	Uma reimpressão ou cancelamento foi executada dentro de uma sessão multi-cartões

RESULTADO_CAPPTA_2	O CapptaGpPlus esta sendo inicializado, tente novamente em alguns instantes

RESULTADO_CAPPTA_3	O formato da requisição recebida pelo CapptaGpPlus é inválido

RESULTADO_CAPPTA_4	Operação cancelada pelo operador

RESULTADO_CAPPTA_5	Pagamento não autorizado/pendente/não encontrado

RESULTADO_CAPPTA_6	Pagamento ou cancelamento negados pela rede adquirente

RESULTADO_CAPPTA_7	Ocorreu um erro interno no CapptaGpPlus

RESULTADO_CAPPTA_8	Ocorreu um erro na comunicação entre a CappAPI e o CapptaGpPlus

RESULTADO_CAPPTA_9	Não é possível realizar uma operação sem que se tenha finalizado o último pagamento

<h1>Etapa 3</h1>

Tempo estimado de 01:00 hora

Conforme mencionado acima a Iteração Tef é muito importante para o perfeito funcionamento da integração, toda as ações de venda e administrativas passam por esta função. 

```javascript
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

```

Dentro de IterarOperacaoTef() temos alguns métodos:

<h3>Requisitar Parametros</h3>


```javascript

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
```


<h3>Resolver Transacao Pendente</h3>

```javascript
string input = Microsoft.VisualBasic.Interaction.InputBox(requisicaoParametros.Mensagem + Environment.NewLine + Environment.NewLine);
	this.cliente.EnviarParametro(input, String.IsNullOrWhiteSpace(input) ? 2 : 1);
```
<h3>Exibir Dados Operacao Aprovada</h3>

```javascript

Private Sub ExibirDadosOperacaoAprovada(resposta As IRespostaOperacaoAprovada)

        Dim mensagemAprovada As New StringBuilder()

        If String.IsNullOrEmpty(resposta.CupomCliente) = False Then
            mensagemAprovada.AppendLine(resposta.CupomCliente.Replace(\, String.Empty)).AppendLine().AppendLine()
        End If

        If String.IsNullOrEmpty(resposta.CupomLojista) = False Then
            mensagemAprovada.Append(resposta.CupomLojista.Replace(\, String.Empty)).AppendLine()
        End If

        If String.IsNullOrEmpty(resposta.CupomReduzido) = False Then
            mensagemAprovada.Append(resposta.CupomReduzido.Replace(\, String.Empty)).AppendLine()
        End If

        AtualizarResultado(mensagemAprovada.ToString())

    End Sub
Obs: no local de barra subistitua por "\"
```

<h3>Finalizar Pagamento</h3>

```javascript
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
```

<h1>Etapa 4</h1>

Tempo estimado de 01:00 hora

Parabéns agora falta pouco, lembrando que a qualquer momento você pode entrar em contato com a equipe tecnica.

Tel: (11) 4302-6179.

Por se tratar de um ambiente de testes, pode ser utilizado cartões reais para as transações, não sera cobrado nada em sua fatura. Se precisar pode utilizar os cartões presentes em nosso [roteiro de teste](http://docs.desktop.cappta.com.br/docs/portf%C3%B3lio-de-cart%C3%B5es-de-testes). Lembrando que vendas digitadas é permitido apenas para a modalidade crédito.

Vamos para a elaboração dos metodos para pagamento.

O primeiro é pagamento débito, o mais simples.

```javascript
 Private Sub ExecutarDebito_Click(sender As Object, e As EventArgs) Handles ExecutarDebito.Click

        If DeveIniciarMultiCartoes() Then
            IniciarMulticartoes()
        End If

        Dim valor As Decimal = NumericUpDownValorPagamentoDebito.Value

        If DeveIniciarMultiCartoes() Then
            IniciarMulticartoes()
        End If

        Dim resultado As Int32
        resultado = cliente.PagamentoDebito(valor)

        If resultado <> 0 Then
            CriarMensagemErroPainel(resultado)
            Return
        End If

        processandoPagamento = True
        IterarOperacaoTef()
    End Sub
```
<h3>Agora pagamento credito:</h3>

```javascript
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
```

<h3>Crediário </h3>

```javascript
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
```

<h1>Etapa 5 </h1>

Tempo estimado de 01:00 hora

**Funções administrativas**

Agora que tratamos as formas de pagamento, podemos partir para as funções administrativas. 

Clientes com frequência pedem a reimpressão de um comprovante ou um cancelamento, as funções administrativas tem a função de deixar praticas e acessiveis estas funções.

<h3>Para reimpressão </h3>
Temos as seguintes formas: 

*Reimpressão por número de controle

*Reimpressão cupom lojista

*Reimpressão cupom cliente

*Reimpressão de todas as vias

```javascript
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

```

<h3>Para Cancelamento </h3>

Para cancelar uma transação é preciso do número de controle e da senha administrativa, esta senha é configurável no Pinpad e por padrão é: **cappta**.  O número de controle é informado na resposta da operação aprovada.

```javascript
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
```
<h1> Etapa 6 </h1>

Tempo estimado de 00:40 minutos

Agora que ja fizemos 80% da integração precisamos trabalhar no Multicartões.

Multicartões ou MultiTef é uma forma de passar mais de um cartão em uma transação, nossa forma de realizar esta tarefa é diferente, se cancelarmos uma venda no meio de uma transação multtef todas são canceladas.

```javascript
 Private Sub IniciarMulticartoes()
        quantidadeCartoes = NumericUpDownQuantidadeDePagamentosMultiTef.Value
        sessaoMultiTefEmAndamento = True
        cliente.IniciarMultiCartoes(quantidadeCartoes)
    End Sub

```
<h6>
Para o código completo basta clonar o repositório, qualquer dúvida entre em contato com o time de homologação e parceria Cappta.
Quando completar a integração basta acessar nossa documentação e seguir os passos do nosso [roteiro](http://docs.desktop.cappta.com.br/docs). </h6>


**Configurando e usando:**

------------------------------------------------------------

- Instale e execute o CapptaGpPlus.exe com os dados forneceidos pela equipe;

- Execute o CapptaGpPlus;

- Extraia e abra o diretório Cappta.Gp.API.VBNet.Samples;

- Abra o arquivo Cappta.Gp.Api.Com.Sample.exe.config (Samples\Binaries\CSharp) em um editor de texto e configure os parametros "Cnpj" e "Pdv" com os dados fornedidos para instalação do CapptaGpPlus (não alterar a Chave de Autenticação); 
**Ex.:** authenticationKey: 'AAAAAAAAAAAAAAAA00000000000000A0', merchantCnpj: '00000000000000', checkoutNumber: 14

- Execute o Cappta.Gp.Api.Com.Sample.exe ou use o código do projeto para fazer as transações de testes.
