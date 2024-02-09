using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EBM_RPA_SipCodeCloudPABX
{
    public partial class FormRPATelefoniaSipCode : Form
    {
        public bool FirstRun { get; set; } = true;
        public string UrlBase = "https://ebmsimulacaofinanciamento.codefi.com.br"; //"http://ebmsimulacaofinanciamento.codefi.com.br"; //"http://localhost:50878"

        public FormRPATelefoniaSipCode()
        {
            InitializeComponent();
            InitializeAsync();
        }

        async void InitializeAsync()
        {
            await webView.EnsureCoreWebView2Async(null);
            webView.CoreWebView2.WebMessageReceived += ReceiveMessage;

            //await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync("window.chrome.webview.postMessage(window.document.URL);");
            //await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync("window.chrome.webview.addEventListener(\'message\', event => alert(event.data));");
        }

        void ReceiveMessage(object sender, CoreWebView2WebMessageReceivedEventArgs args)
        {
            if (DateTime.Now.Hour > 18 || DateTime.Now.Hour < 8)
                return;

            String messageReceived = args.TryGetWebMessageAsString();

            if (messageReceived.StartsWith("http"))
            {
                string url = messageReceived;
                Debug.WriteLine(url);

                // Crie uma instância do HttpClient
                using (HttpClient httpClient = new HttpClient())
                {
                    try
                    {
                        // Faça a requisição GET
                        HttpResponseMessage response = httpClient.GetAsync(url).Result;

                        // Verifique se a resposta foi bem-sucedida
                        if (response.IsSuccessStatusCode)
                        {
                            // Leia o conteúdo da resposta como uma string
                            string conteudo = response.Content.ReadAsStringAsync().Result;
                            Debug.WriteLine("Resposta do servidor:");
                            Debug.WriteLine(conteudo);
                        }
                        else
                        {
                            Debug.WriteLine("Erro na requisição: " + response.StatusCode);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Erro: " + ex.Message);
                    }
                }
            }
            else
            {
                //processa fila
                try
                {
                    var url = $"{UrlBase}/json/AtualizaDadosDashTelefoniaTempoReal";
                    var corpo = messageReceived.Substring(messageReceived.IndexOf("<h2>"), messageReceived.LastIndexOf("Server Time") - messageReceived.IndexOf("<h2>"));

                    // Crie uma instância do HttpClient
                    using (HttpClient httpClient = new HttpClient())
                    {
                        try
                        {
                            // Dados a serem enviados no corpo da requisição (no formato JSON)
                            string jsonContent = corpo;

                            //// Configure o cabeçalho Content-Type para JSON
                            //httpClient.DefaultRequestHeaders.Add("Content-Type", "application/json");

                            // Faça a requisição POST
                            HttpResponseMessage response = httpClient.PostAsync(url, new StringContent(jsonContent, Encoding.UTF8, "application/json")).Result;

                            // Verifique se a resposta foi bem-sucedida
                            if (response.IsSuccessStatusCode)
                            {
                                // Leia o conteúdo da resposta como uma string
                                string conteudo = response.Content.ReadAsStringAsync().Result;
                                Console.WriteLine("Resposta do servidor:");
                                Console.WriteLine(conteudo);
                            }
                            else
                            {
                                Console.WriteLine("Erro na requisição: " + response.StatusCode);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Erro: " + ex.Message);
                        }
                    }
                }
                catch
                {
                    
                }
            }
        }

        private void webView_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            if (FirstRun)
            {
                FirstRun = false;
                SendKeys.Send("usupainel");
                SendKeys.Send("{TAB}");
                SendKeys.Send("acesspnl");
                SendKeys.Send("{ENTER}");
                Thread.Sleep(1000);
                webView.Source = new Uri("http://ebmincorp.sipcode.com.br/queue-stats/index.php");
            }
        }

        private void webView_NavigationStarting(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs e)
        {
            if (FirstRun)
                return;
            else
            {
                webView.CoreWebView2.ExecuteScriptAsync(Script);
                //webView.CoreWebView2.ExecuteScriptAsync($"alert('is not safe, try an https link')");
            }

        }

        public string Script => @"
// Função para interceptar solicitações
function interceptXHR() {
  // Salve a referência ao construtor XMLHttpRequest original
  var originalXHR = window.XMLHttpRequest;

  // Substitua o construtor global XMLHttpRequest
  window.XMLHttpRequest = function () {
    var xhr = new originalXHR();
    
    // Armazenar informações da solicitação
    var requestInfo = {
      method: null,
      url: null,
      params: null,
      headers: null
    };

    // Intercepte eventos, como 'readystatechange'
    xhr.addEventListener('readystatechange', function () {
      if (xhr.readyState === 1) {
        // Quando a solicitação é aberta, capture as informações relevantes
        requestInfo.method = xhr._method || xhr._originalMethod;
        requestInfo.url = xhr._url || xhr._originalURL;
        requestInfo.params = xhr._params || null;
        requestInfo.headers = xhr._headers || null;
      }

      // Faça o que você precisa com a solicitação aqui
      if (xhr.readyState === 4 && xhr.status === 200) {
        console.log('Solicitação XMLHttpRequest interceptada:');
        console.log('Método:', requestInfo.method);
        console.log('URL:', requestInfo.url);
        console.log('Parâmetros:', requestInfo.params);
        console.log('Cabeçalhos:', requestInfo.headers);
        console.log('Resposta:', xhr.responseText);

        let tabelas = document.getElementsByTagName(""table"");
        var atendentesMap = new Map();

        for (let index = 0; index < tabelas.length; index++) {
            var tabela = tabelas[index];
            
            if(tabela.children[0].children[0].children[1].innerText == ""Agente"")
            {
                for (let j = 0; j < tabela.children[1].children.length; j++) {
                    if(atendentesMap.get(tabela.children[1].children[j].children[1].innerText) == undefined){
                        atendentesMap.set(tabela.children[1].children[j].children[1].innerText.trim().replaceAll("" "", """"), 
                        tabela.children[1].children[j].children[2].innerText.trim());
                    }
                }
            }
        }

        var contagemPorValor = {};

        atendentesMap.forEach(function(valor, chave) {
            if (!contagemPorValor[valor.trim().replaceAll("" "", """")]) {
                contagemPorValor[valor.trim().replaceAll("" "", """")] = 1;
            } else {
                contagemPorValor[valor.trim().replaceAll("" "", """")]++;
            }
        });

        if(contagemPorValor[""Disponível""] == undefined)
            contagemPorValor[""Disponível""] = 0;
        
        if(contagemPorValor[""Pausada""] == undefined)
           contagemPorValor[""Pausada""] = 0;

        if(contagemPorValor[""inuse""] == undefined)
            contagemPorValor[""inuse""] = 0;

        if(contagemPorValor[""Indisponível""] == undefined)
            contagemPorValor[""Indisponível""] = 0;

        if(contagemPorValor[""ringing""] == undefined)
            contagemPorValor[""ringing""] = 0;

        for (var valor in contagemPorValor) {
            if (contagemPorValor.hasOwnProperty(valor)) {
                //console.log('chave: ' + valor);
                //console.log('valor: ' + contagemPorValor[valor]);

                var url = '" + UrlBase + @"/json/AtualizaParametros?' +
                'chave=telefoniaRamais' + valor.trim().replaceAll("" "", """") +
                '&valor=' + contagemPorValor[valor];
                
                window.chrome.webview.postMessage(url);
            }
        }

        var tabelaEspera = document.getElementById(""table1"");

        if(tabelaEspera == null)
        {
            var url = '" + UrlBase + @"/json/AtualizaParametros?' +
                'chave=telefoniaRamaisTabelaEspera' +
                '&valor=%20';

            window.chrome.webview.postMessage(url);
            return;
        }

        //45550 - Não Identificado
        //45551 - Não Encontrado
        //45552 - Encontrado
        //45553 - Encontrado VIP
        //994554 - Fila Teste SCBR

       var retornoTabelaEspera = encodeURIComponent(table1.getElementsByTagName(""tbody"")[0].innerHTML.replaceAll(""Callerid"", ""Número"").replaceAll(""45550"", ""Não Identificado"").replaceAll(""45551"", ""Não Encontrado"").replaceAll(""45552"", ""Encontrado"").replaceAll(""45553"", ""Encontrado VIP"").replaceAll(""994554"", ""Fila Teste SCBR"").replaceAll(""\n"", """"));

        var url = '" + UrlBase + @"/json/AtualizaParametros?' +
                'chave=telefoniaRamaisTabelaEspera' +
                '&valor=' + retornoTabelaEspera;

        window.chrome.webview.postMessage(url);

        //window.chrome.webview.postMessage(xhr.responseText);
      }
    });

    // Adicione um método personalizado para definir informações de solicitação
    xhr.setRequestInfo = function (method, url, params, headers) {
      xhr._method = method;
      xhr._url = url;
      xhr._params = params;
      xhr._headers = headers;
    };

    // Substitua os métodos open e send para armazenar informações de solicitação
    var originalOpen = xhr.open;
    xhr.open = function (method, url) {
      this.setRequestInfo(method, url);
      originalOpen.apply(this, arguments);
    };

    var originalSend = xhr.send;
    xhr.send = function (data) {
      this.setRequestInfo(this._method, this._url, data, this.getAllResponseHeaders());
      originalSend.apply(this, arguments);
    };

    return xhr;
  };
}

function reiniciar(){
    window.location.href = ""http://ebmincorp.sipcode.com.br/queue-stats/index.php"";
}

function executar(){
  //Aba de seleção
  if(window. location. href.includes(""index"")){
      //Clica no botão de selecionar todos
      List_move_around('right', true,'queues');

      //Clica no botão ""Mostrar histórico"", que redireciona para a aba ""Atendidas""
      envia();
  }

  //Aba ""Atendidas""
  if(window. location. href.includes(""answered"")){
      //Atendidas Com 15 seg
      var percentual15s = table2.children[1].children[0].children[3].innerText.replace(' %', '');
      var tempoMedioChamadas = document.getElementsByTagName(""table"")[2].children[1].children[2].children[1].innerText.replaceAll("" seg"", """");
      var tempoTotalChamadas = document.getElementsByTagName(""table"")[2].children[1].children[3].children[1].innerText.replaceAll("" min"", """");
      var tempoMedioEspera = document.getElementsByTagName(""table"")[2].children[1].children[4].children[1].innerText.replaceAll("" seg"", """");

      var url = '" + UrlBase + @"/json/TelefoniaAtualizaMetricas?' +
                'percentual15s=' + percentual15s +
                '&tempoMedioChamadas=' + tempoMedioChamadas +              
                '&tempoTotalChamadas=' + tempoTotalChamadas +
                '&tempoMedioEspera=' + tempoMedioEspera;

      window.chrome.webview.postMessage(url);

      window.location.href = ""http://ebmincorp.sipcode.com.br/queue-stats/distribution.php"";
  }

  //Aba ""Distribuição""
  if(window. location. href.includes(""distribution"")){
      var chamadasAtendidas = parseInt(document.getElementsByTagName(""table"")[2].children[1].children[0].children[1].innerText.replaceAll(""chamadas"", """"));
      var chamadasPerdidas = parseInt(document.getElementsByTagName(""table"")[2].children[1].children[1].children[1].innerText.replaceAll(""chamadas"", """"));

      if (isNaN(chamadasAtendidas)) {
        chamadasAtendidas = 0;
      }

      if (isNaN(chamadasPerdidas)) {
        chamadasPerdidas = 0;
      }

      var total = chamadasAtendidas + chamadasPerdidas;

      // Crie a URL com os parâmetros
      var url = '" + UrlBase + @"/json/TelefoniaAtualizaChamadas?' +
                'chamadasAtendidas=' + chamadasAtendidas +
                '&chamadasPerdidas=' + chamadasPerdidas +
                '&total=' + total;

      window.chrome.webview.postMessage(url);

      window.location.href = ""http://ebmincorp.sipcode.com.br/queue-stats/realtime.php"";
  }


  //Aba de tempo real
  if(window. location. href.includes(""realtime"")){
      // Chame a função para interceptar solicitações
      interceptXHR();
  }
}

setTimeout(executar, 2000);

setTimeout(reiniciar, 60000);
";
    }
}
