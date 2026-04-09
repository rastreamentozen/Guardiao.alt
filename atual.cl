/**
 * 🚀 SISTEMA HINOVA + CRM (ESTEIRA DE VERIFICAÇÃO DE SINAL)
 * VERSÃO 27 COLUNAS - MOTOR SGA + PIPELINE DE CRM
 * Arquitetura em Lotes, Defensiva e com Gatilhos Dinâmicos.
 * Código formatado por extenso para máxima legibilidade e manutenção.
 */

// =========================================================================
// ⚙️ BLOCO 1: CONFIGURAÇÕES GLOBAIS DA API E DA PLANILHA
// =========================================================================

var SGA_CONFIG = {
  URL_AUTH: "https://api.hinova.com.br/api/sga/v2/usuario/autenticar",
  URL_CONSULTA_BASE: "https://api.hinova.com.br/api/sga/v2/veiculo/buscar/",
  TOKEN_ASSOCIACAO: "041e6d561c08d16fce2a5beead2ca02fa4f4ee113d51a6f56e9c9e7c89694e4187b271c6cf997f4ece7bf4180e13ee750cb67ffbaa15d5bc82260c3b57a27cfc114322ab6e29a6dc368f1c2b4ea59702456d6c9df528c97f3386aee5978276f6",
  USUARIO: "victor rodrigues",
  SENHA: "ZEN0102"
};

var CONFIG = {
  NOME_ABA_PRINCIPAL: "1 - Verificação",
  ABA_CONCLUIDOS: "Concluídos",
  ABA_SEM_RETORNO: "Sem retorno",
  ABA_ERRO: "Erro",
  ABA_AUDITORIA: "Auditoria",
  
  COL_VERIFICADOR: 0,        // A
  COL_DATA_ENTRADA: 1,       // B
  COL_NOME: 2,               // C
  COL_PLACA: 3,              // D
  COL_CHASSI: 4,             // E
  COL_FIPE: 5,               // F
  COL_EMAIL: 6,              // G
  COL_TELEFONE: 7,           // H
  COL_BAIRRO: 8,             // I
  COL_CIDADE: 9,             // J
  COL_ESTADO: 10,            // K
  COL_CLASSIFICACAO: 11,     // L
  COL_SITUACAO: 12,          // M
  COL_ACAO_API: 13,          // N
  COL_GUARDAO_1: 14,         // O
  COL_GUARDAO_2: 15,         // P
  COL_GUARDAO_3: 16,         // Q
  COL_ENVIAR: 17,            // R
  COL_CHECK_EMAIL: 18,       // S
  COL_CHECK_WHATS: 19,       // T
  COL_DATA_EMAIL: 20,        // U
  COL_DATA_WHATS: 21,        // V
  COL_RESPONSAVEL_EMAIL: 22, // W
  COL_RESPONSAVEL_WHATS: 23, // X
  COL_RESPONDEU: 24,         // Y
  COL_TRAVA_ESTADO: 25,      // Z
  COL_FIPE_BAIXA: 26,        // AA
  
  QTD_COLUNAS: 27,
  LINHAS_CABECHALHO: 1
};

// =========================================================================
// 📱 BLOCO 2: MENUS CASCATA E GATILHO AUTOMÁTICO (ONEDIT)
// =========================================================================

function onOpen() {
  var interfaceUsuario = SpreadsheetApp.getUi();
  var menuPrincipal = interfaceUsuario.createMenu('🚀 Automação Hinova');

  // 1: Dados
  var menuDados = interfaceUsuario.createMenu('📁 1: Dados');
  menuDados.addItem('1 - Processar linhas', 'processarVeiculosSGA');
  menuDados.addItem('2 - Classificar envios', 'classificarEnviosPorStatusSGA');
  menuDados.addItem('3 - Corrigir zero km', 'corrigirZeroKmGuardiao');
  menuDados.addItem('4 - Buscar imei', 'buscarDadosGuardiao');
  menuDados.addItem('5 - Marcar fipe baixa', 'marcarFipeBaixaNativo');
  menuDados.addItem('6 - Sincronizar Notas (Dados Ocultos)', 'sincronizarComentariosOcultos');
  menuDados.addItem('7 - Atualizar Estados e Travar', 'verificarEstadoNativo');
  
  // 2: Disparos e comunicação
  var menuEmails = interfaceUsuario.createMenu('2 - Enviar email...');
  menuEmails.addItem('Jéssica', 'enviarComoJessica');
  menuEmails.addItem('Guilherme', 'enviarComoGuilherme');
  menuEmails.addItem('Campos', 'enviarComoCampos');
  menuEmails.addItem('Priscilane', 'enviarComoPriscilane');
  menuEmails.addItem('Victor', 'enviarComoVictor');
  menuEmails.addItem('Marcelle', 'enviarComoMarcelle');
  menuEmails.addItem('Ana Clara', 'enviarComoAnaClara');

  var menuComunicao = interfaceUsuario.createMenu('📁 2: Disparos e Comunicação');
  menuComunicao.addItem('1 - Autenticar perfil (uso manual)', 'autenticarPerfilManual');
  menuComunicao.addSubMenu(menuEmails);
  menuComunicao.addItem('3 - Atualizar Links WhatsApp', 'executarGeradorLinksMenu');
  menuComunicao.addItem('4 - Verificar Respostas de E-mail', 'placeholderColega');
  menuComunicao.addItem('5 - Conciliar Erros do Gmail', 'placeholderColega');

  // 3: Limpeza e Organização
  var menuLimpeza = interfaceUsuario.createMenu('📁 3: Limpeza e Organização');
  menuLimpeza.addItem('1 - Verificar Duplicados (Placa/Chassi)', 'atualizarVerificacaoDeDuplicados');
  menuLimpeza.addItem('2 - Remover Duplicados (Excluir repetidos)', 'removerDuplicadosPlacaChassi');

  // 4: Migrações de etapa
  var menuMigracoes = interfaceUsuario.createMenu('📁 4: Migrações de etapa');
  menuMigracoes.addItem('1 - Migrar sem retorno', 'executarMigracaoSemRetorno');
  menuMigracoes.addItem('2 - Migrar concluidos', 'migrarConcluidos');

  // Construção do Menu Principal
  menuPrincipal.addSubMenu(menuDados);
  menuPrincipal.addSubMenu(menuComunicao);
  menuPrincipal.addSubMenu(menuLimpeza);
  menuPrincipal.addSeparator();
  menuPrincipal.addSubMenu(menuMigracoes);
  menuPrincipal.addToUi();
}

/**
 * Gatilho Automático Invisível (Ouve cliques nos checkboxes nas abas permitidas)
 */
function onEdit(eventoDeEdicao) {
  if (eventoDeEdicao === undefined || eventoDeEdicao === null || eventoDeEdicao.range === undefined || eventoDeEdicao.range === null) {
    return;
  }

  var abaQueFoiEditada = eventoDeEdicao.range.getSheet().getName();
  var listaDeAbasPermitidasParaOGatilho = [CONFIG.NOME_ABA_PRINCIPAL, CONFIG.ABA_CONCLUIDOS];
  
  if (listaDeAbasPermitidasParaOGatilho.indexOf(abaQueFoiEditada) === -1) {
    return;
  }

  var numeroDaLinhaEditada = eventoDeEdicao.range.getRow();
  var numeroDaColunaEditada = eventoDeEdicao.range.getColumn();
  var valorInseridoNaEdicao = eventoDeEdicao.value;

  if (numeroDaLinhaEditada <= CONFIG.LINHAS_CABECHALHO) {
    return;
  }

  if ((numeroDaColunaEditada === 19 || numeroDaColunaEditada === 20) && valorInseridoNaEdicao === "TRUE") {
    
    var objetoAbaAtual = eventoDeEdicao.range.getSheet();
    var dataExataAtual = new Date();
    var stringDataHoraFormatada = Utilities.formatDate(dataExataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    var stringPerfilLogadoNaSessao = PropertiesService.getUserProperties().getProperty("PERFIL_MANUAL");
    
    if (stringPerfilLogadoNaSessao === null || stringPerfilLogadoNaSessao === "") {
      stringPerfilLogadoNaSessao = "Não Autenticado (Manual)";
    }

    if (numeroDaColunaEditada === 19) { 
      
      var celulaDataEmail = objetoAbaAtual.getRange(numeroDaLinhaEditada, CONFIG.COL_DATA_EMAIL + 1);
      celulaDataEmail.setValue(stringDataHoraFormatada);
      
      var celulaResponsavelEmail = objetoAbaAtual.getRange(numeroDaLinhaEditada, CONFIG.COL_RESPONSAVEL_EMAIL + 1);
      celulaResponsavelEmail.setValue(stringPerfilLogadoNaSessao);
      
      var arrayComDadosDaLinhaDoEmail = objetoAbaAtual.getRange(numeroDaLinhaEditada, 1, 1, CONFIG.QTD_COLUNAS).getValues();
      registrarAuditoriaLog(arrayComDadosDaLinhaDoEmail, "E-mail Marcado Manualmente", stringDataHoraFormatada, stringPerfilLogadoNaSessao);
      
    } else if (numeroDaColunaEditada === 20) { 
      
      var celulaDataWhatsApp = objetoAbaAtual.getRange(numeroDaLinhaEditada, CONFIG.COL_DATA_WHATS + 1);
      celulaDataWhatsApp.setValue(stringDataHoraFormatada);
      
      var celulaResponsavelWhatsApp = objetoAbaAtual.getRange(numeroDaLinhaEditada, CONFIG.COL_RESPONSAVEL_WHATS + 1);
      celulaResponsavelWhatsApp.setValue(stringPerfilLogadoNaSessao);
      
      var arrayComDadosDaLinhaDoWhatsApp = objetoAbaAtual.getRange(numeroDaLinhaEditada, 1, 1, CONFIG.QTD_COLUNAS).getValues();
      registrarAuditoriaLog(arrayComDadosDaLinhaDoWhatsApp, "WhatsApp Marcado Manualmente", stringDataHoraFormatada, stringPerfilLogadoNaSessao);
      
    }
  }
}

function placeholderColega() {
  Browser.msgBox("⏳ Aguardando a inserção do código pelo seu colega!");
}

// =========================================================================
// 🔗 BLOCO 3: INTEGRAÇÃO HINOVA (SGA)
// =========================================================================

function autenticarSGA() {
  try {
    var stringDoPayloadFormatado = JSON.stringify({
      "usuario": SGA_CONFIG.USUARIO,
      "senha": SGA_CONFIG.SENHA
    });

    var opcoesDaRequisicaoAutenticacao = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO,
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      "payload": stringDoPayloadFormatado,
      "muteHttpExceptions": true
    };
    
    var respostaDaRequisicaoAutenticacao = UrlFetchApp.fetch(SGA_CONFIG.URL_AUTH, opcoesDaRequisicaoAutenticacao);
    var codigoHTTPDeRetorno = respostaDaRequisicaoAutenticacao.getResponseCode();
    var textoPuroDeRetorno = respostaDaRequisicaoAutenticacao.getContentText();
    
    if (codigoHTTPDeRetorno !== 200) {
      Browser.msgBox("❌ Falha de Login SGA:\n\nCódigo do Erro: " + codigoHTTPDeRetorno + "\nResposta: " + textoPuroDeRetorno);
      return null;
    }
    
    var objetoJsonParseado = JSON.parse(textoPuroDeRetorno);
    
    if (objetoJsonParseado.token_usuario !== undefined && objetoJsonParseado.token_usuario !== null) {
      return objetoJsonParseado.token_usuario;
    } else {
      return null;
    }
    
  } catch(erroNaConexao) {
    Browser.msgBox("❌ Erro ao conectar com Hinova: " + erroNaConexao.message);
    return null;
  }
}

/**
 * 🔗 PROCESSAMENTO SGA
 * CORREÇÃO: parse da data corrigido — .split(" ")[0] e partesData[0]
 */
function processarVeiculosSGA() {
  var stringTokenDeSessao = autenticarSGA();
  if (!stringTokenDeSessao) {
    Browser.msgBox("❌ Falha na autenticação SGA Hinova.");
    return;
  }

  var arquivoDePlanilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
  var abaDeVerificacaoPrincipal = arquivoDePlanilhaAtiva.getSheetByName(CONFIG.NOME_ABA_PRINCIPAL);
  
  if (!abaDeVerificacaoPrincipal) {
    Browser.msgBox("Aba principal não encontrada.");
    return;
  }

  var numeroDaUltimaLinhaPreenchida = abaDeVerificacaoPrincipal.getLastRow();
  if (numeroDaUltimaLinhaPreenchida <= CONFIG.LINHAS_CABECHALHO) {
    Browser.msgBox("Nenhuma linha de dados encontrada para processar.");
    return; 
  }
  
  var intervaloCompletoDeDados = abaDeVerificacaoPrincipal.getRange(1, 1, numeroDaUltimaLinhaPreenchida, CONFIG.QTD_COLUNAS);
  var matrizMestreDosDados = intervaloCompletoDeDados.getValues();
  var variavelHouveAlgumaAlteracaoNaPlanilha = false;
  var contadorDeVeiculosProcessadosComSucesso = 0;

  var dicionarioDeClassificacao = {
    "1": "Concluido", "2": "Inativo", "3": "Pendente", "4": "Inadimplente", 
    "5": "Negado", "6": "1º boleto pago e s/ rastreador", "7": "Evento", 
    "8": "Indenizado", "9": "Pre-aprovado", "10": "Evento negado", 
    "12": "Inativos com rastreador", "13": "Inativos sem rastreador", 
    "14": "Instalação pendente", "15": "Publicado no spc/serasa", 
    "16": "Aguardando taxa instalação", "17": "Faltando vistoria, sem rastreador", 
    "18": "Pendente", "19": "Aguardando indenização", "20": "Roubo/furto"
  };

  for (var indiceDeLinha = CONFIG.LINHAS_CABECHALHO; indiceDeLinha < matrizMestreDosDados.length; indiceDeLinha++) {
    
    var stringDeAcaoDaAPI = matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ACAO_API] ? matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ACAO_API].toString().trim().toLowerCase() : "";
    if (stringDeAcaoDaAPI !== "verificar") {
      continue; 
    }
    
    var stringDaPlacaDoCliente = matrizMestreDosDados[indiceDeLinha][CONFIG.COL_PLACA] ? matrizMestreDosDados[indiceDeLinha][CONFIG.COL_PLACA].toString().trim().toUpperCase() : "";
    var stringDoChassiDoCliente = matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CHASSI] ? matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CHASSI].toString().trim().toUpperCase() : "";
    
    var valorDeBuscaDefinido = stringDaPlacaDoCliente || stringDoChassiDoCliente;
    var parametroDeBuscaSelecionado = stringDaPlacaDoCliente ? "placa" : "chassi";
    
    if (!valorDeBuscaDefinido) {
      matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ACAO_API] = "Falta Placa/Chassi";
      variavelHouveAlgumaAlteracaoNaPlanilha = true;
      continue;
    }
    
    var stringDoEndpointDeConsulta = SGA_CONFIG.URL_CONSULTA_BASE + encodeURIComponent(valorDeBuscaDefinido) + "/" + parametroDeBuscaSelecionado;
    
    try {
      var opcoesDaRequisicaoConsulta = {
        "method": "GET",
        "headers": { "Authorization": "Bearer " + stringTokenDeSessao },
        "muteHttpExceptions": true 
      };

      var respostaDaAPI = UrlFetchApp.fetch(stringDoEndpointDeConsulta, opcoesDaRequisicaoConsulta);
      
      if (respostaDaAPI.getResponseCode() == 200) {
        var objetoRetornoParseado = JSON.parse(respostaDaAPI.getContentText());
        var veiculoEncontradoNoSGA = Array.isArray(objetoRetornoParseado) ? objetoRetornoParseado[0] : objetoRetornoParseado;
        
        if (veiculoEncontradoNoSGA && veiculoEncontradoNoSGA.codigo_veiculo) {
          
          var stringDataDeEntradaCrua = veiculoEncontradoNoSGA.data_cadastro || veiculoEncontradoNoSGA.data_contrato || veiculoEncontradoNoSGA.data_cadastro_associado || veiculoEncontradoNoSGA.data_contrato_associado || veiculoEncontradoNoSGA.data_inicio_contrato_veiculo || veiculoEncontradoNoSGA.data_inicio_contrato_associado || veiculoEncontradoNoSGA.data_cadastro_veiculo || veiculoEncontradoNoSGA.data_criacao || "";
          
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_DATA_ENTRADA] && stringDataDeEntradaCrua) {
            // ✅ CORREÇÃO: .split(" ")[0] para isolar a data, e partesData[0] para o ano
            var stringData = stringDataDeEntradaCrua.toString().trim().split(" ")[0];
            if (stringData.indexOf("-") > -1) { 
              var partesData = stringData.split("-");
              if (partesData.length === 3) {
                matrizMestreDosDados[indiceDeLinha][CONFIG.COL_DATA_ENTRADA] = partesData[2] + "/" + partesData[1] + "/" + partesData[0];
              }
            } else { 
              matrizMestreDosDados[indiceDeLinha][CONFIG.COL_DATA_ENTRADA] = stringData;
            }
          }
          
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_NOME]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_NOME] = (veiculoEncontradoNoSGA.nome || veiculoEncontradoNoSGA.nome_associado || "").toString().trim();
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_PLACA]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_PLACA] = (veiculoEncontradoNoSGA.placa || "").toString().trim().toUpperCase();
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CHASSI]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CHASSI] = (veiculoEncontradoNoSGA.chassi || "").toString().trim().toUpperCase();
          if (veiculoEncontradoNoSGA.valor_fipe) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_FIPE] = parseFloat(veiculoEncontradoNoSGA.valor_fipe.toString().replace(",", ".")); 
          
          var stringCodigoDeClassificacao = veiculoEncontradoNoSGA.codigo_classificacao ? veiculoEncontradoNoSGA.codigo_classificacao.toString() : "";
          matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CLASSIFICACAO] = dicionarioDeClassificacao[stringCodigoDeClassificacao] || "Código " + stringCodigoDeClassificacao;

          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_EMAIL]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_EMAIL] = (veiculoEncontradoNoSGA.email || "").toString().trim().toLowerCase();

          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_TELEFONE]) {
            var fone = (veiculoEncontradoNoSGA.ddd_celular ? "(" + veiculoEncontradoNoSGA.ddd_celular + ") " : "") + (veiculoEncontradoNoSGA.telefone_celular || veiculoEncontradoNoSGA.telefone || "");
            matrizMestreDosDados[indiceDeLinha][CONFIG.COL_TELEFONE] = fone.trim();
          }
          
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_BAIRRO]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_BAIRRO] = (veiculoEncontradoNoSGA.bairro || "").toString().trim();
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CIDADE]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_CIDADE] = (veiculoEncontradoNoSGA.cidade || "").toString().trim();
          if (!matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ESTADO]) matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ESTADO] = (veiculoEncontradoNoSGA.estado || "").toString().trim();
          
          var stringDoStatusQueFoiRetornado = veiculoEncontradoNoSGA.descricao_situacao_veiculo || veiculoEncontradoNoSGA.situacao_veiculo || veiculoEncontradoNoSGA.descricao_situacao || veiculoEncontradoNoSGA.situacao || "";
          
          if (!stringDoStatusQueFoiRetornado) {
            var stringDoCodigoDaSituacao = veiculoEncontradoNoSGA.codigo_situacao_veiculo || veiculoEncontradoNoSGA.codigo_situacao;
            if (stringDoCodigoDaSituacao == 1 || stringDoCodigoDaSituacao == "1") stringDoStatusQueFoiRetornado = "ATIVO";
            else if (stringDoCodigoDaSituacao == 2 || stringDoCodigoDaSituacao == "2") stringDoStatusQueFoiRetornado = "INATIVO";
            else if (stringDoCodigoDaSituacao == 3 || stringDoCodigoDaSituacao == "3") stringDoStatusQueFoiRetornado = "PENDENTE";
            else if (stringDoCodigoDaSituacao) stringDoStatusQueFoiRetornado = "Cód: " + stringDoCodigoDaSituacao;
          }

          matrizMestreDosDados[indiceDeLinha][CONFIG.COL_SITUACAO] = (stringDoStatusQueFoiRetornado || "S/ STATUS").toString().trim().toUpperCase();
          matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ACAO_API] = "Concluído";
          
          contadorDeVeiculosProcessadosComSucesso++;
          variavelHouveAlgumaAlteracaoNaPlanilha = true;

        } else {
          matrizMestreDosDados[indiceDeLinha][CONFIG.COL_ACAO_API] = "Não Encontrado";
          variavelHouveAlgumaAlteracaoNaPlanilha = true;
        }
      }
    } catch (erroDeExecucaoDaRequisicao) { 
        Logger.log(erroDeExecucaoDaRequisicao.message); 
    }
    
    Utilities.sleep(200);
  }

  if (variavelHouveAlgumaAlteracaoNaPlanilha) {
    intervaloCompletoDeDados.setValues(matrizMestreDosDados);
    abaDeVerificacaoPrincipal.getRange(2, CONFIG.COL_FIPE + 1, numeroDaUltimaLinhaPreenchida - 1, 1).setNumberFormat('R$ #,##0.00');
    SpreadsheetApp.flush(); 
  }

  Browser.msgBox("✅ Verificação Concluída! " + contadorDeVeiculosProcessadosComSucesso + " veículos processados.");
}

// =========================================================================
// 📊 BLOCO 4: CRUZAMENTO DE DADOS (GUARDIÃO)
// =========================================================================

function buscarDadosGuardiao() {
  var arquivoAtivoDaPlanilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaAlvoDestino = arquivoAtivoDaPlanilha.getSheetByName(CONFIG.NOME_ABA_PRINCIPAL);
  var abaReferenciaOrigem = arquivoAtivoDaPlanilha.getSheetByName("exportacao_guardiao");
  var abaReferenciaAtivos = arquivoAtivoDaPlanilha.getSheetByName("ativos");
  
  if (abaAlvoDestino === null) {
      Browser.msgBox("❌ Aba principal não encontrada!");
      return;
  }
  
  var numeroDaUltimaLinhaDoDestino = abaAlvoDestino.getLastRow();
  if (numeroDaUltimaLinhaDoDestino <= CONFIG.LINHAS_CABECHALHO) {
      Browser.msgBox("Nenhum dado para buscar.");
      return;
  }

  var intervaloCompletoDeDestino = abaAlvoDestino.getRange(1, 1, numeroDaUltimaLinhaDoDestino, CONFIG.QTD_COLUNAS);
  var matrizDosDadosDeDestino = intervaloCompletoDeDestino.getValues();
  
  function formatarDadoDeManeiraSegura(valorRecebido) {
    if (valorRecebido === undefined || valorRecebido === null || valorRecebido === "") {
        return "Não informado";
    }
    if (valorRecebido instanceof Date) {
        return Utilities.formatDate(valorRecebido, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    var stringConvertidaETratada = valorRecebido.toString().trim();
    if (stringConvertidaETratada === "") {
        return "Não informado";
    }
    return stringConvertidaETratada;
  }
  
  var dicionarioMapeadoDoGuardiao = {};
  var dicionarioMapeadoDosAtivos = {};
  
  // 1. Dicionário do Guardião
  if (abaReferenciaOrigem !== null) {
    var intervaloGeralDaOrigem = abaReferenciaOrigem.getDataRange();
    var matrizDadosGeraisDaOrigem = intervaloGeralDaOrigem.getValues();
    
    var NUMERO_COLUNA_PLACA_GUARDIAO = 3; 
    var NUMERO_COLUNA_IMEI_GUARDIAO = 6; 
    var NUMERO_COLUNA_GD2_GUARDIAO = 17; 
    var NUMERO_COLUNA_GD3_GUARDIAO = 18; 
    var NUMERO_COLUNA_DATA_GUARDIAO = 20; 

    for (var contadorDeLinhasGuardiao = 1; contadorDeLinhasGuardiao < matrizDadosGeraisDaOrigem.length; contadorDeLinhasGuardiao++) {
      var stringDaPlacaNoGuardiao = "";
      if (matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_PLACA_GUARDIAO] !== undefined && matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_PLACA_GUARDIAO] !== null && matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_PLACA_GUARDIAO] !== "") {
          stringDaPlacaNoGuardiao = matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_PLACA_GUARDIAO].toString().trim().toUpperCase();
      }

      if (stringDaPlacaNoGuardiao !== "") {
        var stringDoImeiExtraido = formatarDadoDeManeiraSegura(matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_IMEI_GUARDIAO]);
        var stringDoValorGuarda2 = formatarDadoDeManeiraSegura(matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_GD2_GUARDIAO]);
        var stringDoValorGuarda3 = formatarDadoDeManeiraSegura(matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_GD3_GUARDIAO]);
        var stringDaDataExtraida = formatarDadoDeManeiraSegura(matrizDadosGeraisDaOrigem[contadorDeLinhasGuardiao][NUMERO_COLUNA_DATA_GUARDIAO]);
        
        if (stringDoImeiExtraido !== "Não informado" || stringDoValorGuarda2 !== "Não informado" || stringDoValorGuarda3 !== "Não informado") {
          if (dicionarioMapeadoDoGuardiao[stringDaPlacaNoGuardiao] === undefined) {
              dicionarioMapeadoDoGuardiao[stringDaPlacaNoGuardiao] = [];
          }
          
          var objetoEmpacotadoRegistroGuardiao = { 
              colunaIMEI: stringDoImeiExtraido, 
              colunaGuarda2: stringDoValorGuarda2, 
              colunaGuarda3: stringDoValorGuarda3, 
              colunaData: stringDaDataExtraida 
          };

          dicionarioMapeadoDoGuardiao[stringDaPlacaNoGuardiao].push(objetoEmpacotadoRegistroGuardiao);
        }
      }
    }
  }
  
  // 2. Dicionário Aba Ativos
  if (abaReferenciaAtivos !== null) {
    var intervaloGeralDosAtivos = abaReferenciaAtivos.getDataRange();
    var matrizDadosGeraisDosAtivos = intervaloGeralDosAtivos.getValues();
    
    var NUMERO_COLUNA_PLACA_ATIVOS = 3; 
    var NUMERO_COLUNA_CHASSI_ATIVOS = 4; 
    var NUMERO_COLUNA_IMEI_ATIVOS = 27; 
    var NUMERO_COLUNA_DATA_ATIVOS = 37; 
    
    for (var contadorDeLinhasAtivos = 1; contadorDeLinhasAtivos < matrizDadosGeraisDosAtivos.length; contadorDeLinhasAtivos++) {
      var stringDaPlacaNaAbaAtivos = "";
      if (matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_PLACA_ATIVOS] !== undefined && matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_PLACA_ATIVOS] !== null && matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_PLACA_ATIVOS] !== "") {
          stringDaPlacaNaAbaAtivos = matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_PLACA_ATIVOS].toString().trim().toUpperCase();
      }

      var stringDoChassiNaAbaAtivos = "";
      if (matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_CHASSI_ATIVOS] !== undefined && matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_CHASSI_ATIVOS] !== null && matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_CHASSI_ATIVOS] !== "") {
          stringDoChassiNaAbaAtivos = matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_CHASSI_ATIVOS].toString().trim().toUpperCase();
      }
      
      var stringDoImeiExtraidoAtivos = formatarDadoDeManeiraSegura(matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_IMEI_ATIVOS]);
      var stringDaDataExtraidaAtivos = formatarDadoDeManeiraSegura(matrizDadosGeraisDosAtivos[contadorDeLinhasAtivos][NUMERO_COLUNA_DATA_ATIVOS]);
      
      if (stringDoImeiExtraidoAtivos !== "Não informado") {
        
        var objetoEmpacotadoRegistroAtivos = { 
            colunaIMEI: stringDoImeiExtraidoAtivos, 
            colunaGuarda2: "Não informado", 
            colunaGuarda3: "Não informado", 
            colunaData: stringDaDataExtraidaAtivos 
        };

        if (stringDaPlacaNaAbaAtivos !== "") {
          if (dicionarioMapeadoDosAtivos[stringDaPlacaNaAbaAtivos] === undefined) {
              dicionarioMapeadoDosAtivos[stringDaPlacaNaAbaAtivos] = [];
          }
          dicionarioMapeadoDosAtivos[stringDaPlacaNaAbaAtivos].push(objetoEmpacotadoRegistroAtivos);
        }

        if (stringDoChassiNaAbaAtivos !== "" && stringDoChassiNaAbaAtivos !== stringDaPlacaNaAbaAtivos) {
          if (dicionarioMapeadoDosAtivos[stringDoChassiNaAbaAtivos] === undefined) {
              dicionarioMapeadoDosAtivos[stringDoChassiNaAbaAtivos] = [];
          }
          dicionarioMapeadoDosAtivos[stringDoChassiNaAbaAtivos].push(objetoEmpacotadoRegistroAtivos);
        }
      }
    }
  }
  
  // 3. Preenchimento
  var dicionarioDeContagemAbaPrincipal = {};
  for (var k = CONFIG.LINHAS_CABECHALHO; k < matrizDosDadosDeDestino.length; k++) {
    
    var placaDestino = "";
    if (matrizDosDadosDeDestino[k][CONFIG.COL_PLACA] !== undefined && matrizDosDadosDeDestino[k][CONFIG.COL_PLACA] !== null) {
        placaDestino = matrizDosDadosDeDestino[k][CONFIG.COL_PLACA].toString().trim().toUpperCase();
    }

    var chassiDestino = "";
    if (matrizDosDadosDeDestino[k][CONFIG.COL_CHASSI] !== undefined && matrizDosDadosDeDestino[k][CONFIG.COL_CHASSI] !== null) {
        chassiDestino = matrizDosDadosDeDestino[k][CONFIG.COL_CHASSI].toString().trim().toUpperCase();
    }

    var identificadorPrincipal = "";
    if (placaDestino !== "") {
        identificadorPrincipal = placaDestino;
    } else {
        identificadorPrincipal = chassiDestino;
    }

    if (identificadorPrincipal !== "") {
        if (dicionarioDeContagemAbaPrincipal[identificadorPrincipal] === undefined) {
            dicionarioDeContagemAbaPrincipal[identificadorPrincipal] = 0;
        }
        dicionarioDeContagemAbaPrincipal[identificadorPrincipal] = dicionarioDeContagemAbaPrincipal[identificadorPrincipal] + 1;
    }
  }
  
  var indiceInicio = CONFIG.LINHAS_CABECHALHO;
  var indiceFim = matrizDosDadosDeDestino.length - 1;
  var contadorAtualizados = 0;
  var houveAlteracaoNoCruzamento = false;
  var usoSequencialMemoria = {};
  
  for (var m = indiceInicio; m <= indiceFim; m++) {
    
    var placaBusca = "";
    if (matrizDosDadosDeDestino[m][CONFIG.COL_PLACA] !== undefined && matrizDosDadosDeDestino[m][CONFIG.COL_PLACA] !== null) {
        placaBusca = matrizDosDadosDeDestino[m][CONFIG.COL_PLACA].toString().trim().toUpperCase();
    }

    var chassiBusca = "";
    if (matrizDosDadosDeDestino[m][CONFIG.COL_CHASSI] !== undefined && matrizDosDadosDeDestino[m][CONFIG.COL_CHASSI] !== null) {
        chassiBusca = matrizDosDadosDeDestino[m][CONFIG.COL_CHASSI].toString().trim().toUpperCase();
    }

    var idCrucial = "";
    if (placaBusca !== "") {
        idCrucial = placaBusca;
    } else {
        idCrucial = chassiBusca;
    }
    
    if (idCrucial !== "") {
      
      var registrosGuardiao = [];
      if (placaBusca !== "" && dicionarioMapeadoDoGuardiao[placaBusca] !== undefined) {
          registrosGuardiao = dicionarioMapeadoDoGuardiao[placaBusca];
      } else if (chassiBusca !== "" && dicionarioMapeadoDoGuardiao[chassiBusca] !== undefined) {
          registrosGuardiao = dicionarioMapeadoDoGuardiao[chassiBusca];
      }

      var registrosAtivos = [];
      if (placaBusca !== "" && dicionarioMapeadoDosAtivos[placaBusca] !== undefined) {
          registrosAtivos = dicionarioMapeadoDosAtivos[placaBusca];
      } else if (chassiBusca !== "" && dicionarioMapeadoDosAtivos[chassiBusca] !== undefined) {
          registrosAtivos = dicionarioMapeadoDosAtivos[chassiBusca];
      }
      
      var arrayUnificadoRegistros = [];
      var imeiJaAdicionados = {};
      
      for(var g = 0; g < registrosGuardiao.length; g++) { 
          arrayUnificadoRegistros.push(registrosGuardiao[g]); 
          var valorImeiG = registrosGuardiao[g].colunaIMEI;
          imeiJaAdicionados[valorImeiG] = true; 
      }

      for(var a = 0; a < registrosAtivos.length; a++) {
        var valorImeiA = registrosAtivos[a].colunaIMEI;
        if(imeiJaAdicionados[valorImeiA] === undefined) { 
            arrayUnificadoRegistros.push(registrosAtivos[a]); 
            imeiJaAdicionados[valorImeiA] = true; 
        }
      }

      if (arrayUnificadoRegistros.length > 0) {
        
        if (dicionarioDeContagemAbaPrincipal[idCrucial] === 1) {
          
          var stringDatasUnificadas = arrayUnificadoRegistros.map(function(registroTemp){ return registroTemp.colunaData; }).join("\n");
          var stringImeisUnificados = arrayUnificadoRegistros.map(function(registroTemp){ return registroTemp.colunaIMEI; }).join("\n");
          var stringGuarda2Unificados = arrayUnificadoRegistros.map(function(registroTemp){ return registroTemp.colunaGuarda2; }).join("\n");
          var stringGuarda3Unificados = arrayUnificadoRegistros.map(function(registroTemp){ return registroTemp.colunaGuarda3; }).join("\n");

          matrizDosDadosDeDestino[m][CONFIG.COL_DATA_ENTRADA] = stringDatasUnificadas;
          matrizDosDadosDeDestino[m][CONFIG.COL_GUARDAO_1] = stringImeisUnificados;
          matrizDosDadosDeDestino[m][CONFIG.COL_GUARDAO_2] = stringGuarda2Unificados;
          matrizDosDadosDeDestino[m][CONFIG.COL_GUARDAO_3] = stringGuarda3Unificados;

        } else {
          
          var indiceDoVeiculoRepetido = 0;
          if (usoSequencialMemoria[idCrucial] !== undefined) {
              indiceDoVeiculoRepetido = usoSequencialMemoria[idCrucial];
          }
          
          if (indiceDoVeiculoRepetido < arrayUnificadoRegistros.length) {
            
            var registroUnicoFocado = arrayUnificadoRegistros[indiceDoVeiculoRepetido];

            matrizDosDadosDeDestino[m][CONFIG.COL_DATA_ENTRADA] = registroUnicoFocado.colunaData;
            matrizDosDadosDeDestino[m][CONFIG.COL_GUARDAO_1] = registroUnicoFocado.colunaIMEI;
            matrizDosDadosDeDestino[m][CONFIG.COL_GUARDAO_2] = registroUnicoFocado.colunaGuarda2;
            matrizDosDadosDeDestino[m][CONFIG.COL_GUARDAO_3] = registroUnicoFocado.colunaGuarda3;
            
            usoSequencialMemoria[idCrucial] = indiceDoVeiculoRepetido + 1;
          }
        }

        houveAlteracaoNoCruzamento = true;
        contadorAtualizados++;
      }
    }
  }
  
  if (houveAlteracaoNoCruzamento === true) {
    var intervaloParaGravar = abaAlvoDestino.getRange(2, 1, matrizDosDadosDeDestino.length - 1, CONFIG.QTD_COLUNAS);
    var dadosParaGravarSemCabecalho = matrizDosDadosDeDestino.slice(1);
    
    intervaloParaGravar.setValues(dadosParaGravarSemCabecalho);
    SpreadsheetApp.flush();
    
    Browser.msgBox("✅ " + contadorAtualizados + " veículos atualizados com os dados de Guardião.");
  } else {
    Browser.msgBox("⚠️ Nenhum dado novo encontrado.");
  }
}

function corrigirZeroKmGuardiao() {
  var planilhaMestre = SpreadsheetApp.getActiveSpreadsheet();
  var abaGuardiao = planilhaMestre.getSheetByName("exportacao_guardiao");
  var abaAtivos = planilhaMestre.getSheetByName("ativos");
  
  if (abaGuardiao === null || abaAtivos === null) {
      Browser.msgBox("Erro: Abas exportacao_guardiao ou ativos não encontradas!");
      return;
  }
  
  var ultimaLinhaAbaAtivos = abaAtivos.getLastRow();
  if (ultimaLinhaAbaAtivos < 1) {
      return;
  }

  var intervaloAtivos = abaAtivos.getRange(1, 1, ultimaLinhaAbaAtivos, 40);
  var dadosAtivosExtraidos = intervaloAtivos.getValues();
  
  var mapaDeImeisEChassis = new Map();

  for (var i = 1; i < dadosAtivosExtraidos.length; i++) {
    
    var stringImeiNaAbaAtivos = "";
    if (dadosAtivosExtraidos[i][27] !== undefined && dadosAtivosExtraidos[i][27] !== null) {
        stringImeiNaAbaAtivos = dadosAtivosExtraidos[i][27].toString().trim();
    }

    var stringChassiNaAbaAtivos = "";
    if (dadosAtivosExtraidos[i][4] !== undefined && dadosAtivosExtraidos[i][4] !== null) {
        stringChassiNaAbaAtivos = dadosAtivosExtraidos[i][4].toString().trim().toUpperCase();
    }

    if (stringImeiNaAbaAtivos !== "" && stringChassiNaAbaAtivos !== "") {
        mapaDeImeisEChassis.set(stringImeiNaAbaAtivos, stringChassiNaAbaAtivos);
    }
  }
  
  var ultimaLinhaAbaGuardiao = abaGuardiao.getLastRow();
  if (ultimaLinhaAbaGuardiao < 1) {
      return;
  }

  var intervaloGuardiao = abaGuardiao.getRange(1, 1, ultimaLinhaAbaGuardiao, 20);
  var dadosGuardiaoExtraidos = intervaloGuardiao.getValues();
  
  var alterouZeroKm = false;
  var contagemCorrigidos = 0;
  
  for (var j = 1; j < dadosGuardiaoExtraidos.length; j++) {
    
    var stringPlacaNoGuardiao = "";
    if (dadosGuardiaoExtraidos[j][3] !== undefined && dadosGuardiaoExtraidos[j][3] !== null) {
        stringPlacaNoGuardiao = dadosGuardiaoExtraidos[j][3].toString().trim();
    }

    var stringImeiNoGuardiao = "";
    if (dadosGuardiaoExtraidos[j][6] !== undefined && dadosGuardiaoExtraidos[j][6] !== null) {
        stringImeiNoGuardiao = dadosGuardiaoExtraidos[j][6].toString().trim();
    }
    
    var placaEmMinusculo = stringPlacaNoGuardiao.toLowerCase();

    if (placaEmMinusculo === "zero km" || placaEmMinusculo === "zerokm") {
      
      if (stringImeiNoGuardiao !== "" && mapaDeImeisEChassis.has(stringImeiNoGuardiao) === true) {
        var chassiMapeadoEncontrado = mapaDeImeisEChassis.get(stringImeiNoGuardiao);
        dadosGuardiaoExtraidos[j][3] = chassiMapeadoEncontrado;
        
        alterouZeroKm = true;
        contagemCorrigidos++;
      }

    }
  }
  
  if (alterouZeroKm === true) {
    intervaloGuardiao.setValues(dadosGuardiaoExtraidos);
    SpreadsheetApp.flush();
    Browser.msgBox("✅ Sucesso! " + contagemCorrigidos + " veículos Zero Km foram atualizados para o Chassi.");
  } else {
    Browser.msgBox("Aviso: Nenhum 'Zero KM' precisava ser substituído no Guardião.");
  }
}

// =========================================================================
// 💬 BLOCO 5: COMUNICAÇÃO & DISPAROS (E-MAIL E WHATSAPP)
// =========================================================================

function autenticarPerfilManual() {
  var interfaceGraficaDoUsuario = SpreadsheetApp.getUi();
  var caixaDeRespostaAberta = interfaceGraficaDoUsuario.prompt("👤 Autenticar Perfil", "Digite o seu nome (Ex: Guilherme, Jéssica):", interfaceGraficaDoUsuario.ButtonSet.OK_CANCEL);
  
  if (caixaDeRespostaAberta.getSelectedButton() === interfaceGraficaDoUsuario.Button.OK) {
    var stringNomeDigitado = caixaDeRespostaAberta.getResponseText().trim();
    if (stringNomeDigitado !== "") {
      PropertiesService.getUserProperties().setProperty("PERFIL_MANUAL", stringNomeDigitado);
      Browser.msgBox("✅ Autenticado com sucesso! Perfil ativo: " + stringNomeDigitado);
    }
  }
}

function enviarComoJessica() { dispararFluxo("Jéssica"); }
function enviarComoGuilherme() { dispararFluxo("Guilherme"); }
function enviarComoCampos() { dispararFluxo("Campos"); }
function enviarComoPriscilane() { dispararFluxo("Priscilane"); }
function enviarComoVictor() { dispararFluxo("Victor"); }
function enviarComoMarcelle() { dispararFluxo("Marcelle"); }
function enviarComoAnaClara() { dispararFluxo("Ana Clara"); }

function dispararFluxo(nomeDoResponsavelRecebido) {
  var abaPlanilhaAtualNaSessao = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var stringNomeDaAbaAtualNaSessao = abaPlanilhaAtualNaSessao.getName();
  
  if (stringNomeDaAbaAtualNaSessao !== CONFIG.NOME_ABA_PRINCIPAL && stringNomeDaAbaAtualNaSessao !== CONFIG.ABA_CONCLUIDOS && stringNomeDaAbaAtualNaSessao !== "2 - Expirado") {
    Browser.msgBox("⚠️ E-mails não podem ser enviados a partir desta aba.");
    return;
  }

  var aAbaAtualEhParaExpirados = false;
  if (stringNomeDaAbaAtualNaSessao.indexOf("2 -") > -1 || stringNomeDaAbaAtualNaSessao.toLowerCase().indexOf("expirado") > -1) {
      aAbaAtualEhParaExpirados = true;
  }

  var matrizCompletaDeDadosDaAbaAtual = abaPlanilhaAtualNaSessao.getDataRange().getValues();
  var regexPadraoDeFormatoDeEmail = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  
  var dataExataAgora = new Date();
  var stringDaDataFormatadaParaRegistro = Utilities.formatDate(dataExataAgora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  
  var totalDeEnviosComSucesso = 0;
  var totalDeErrosNaTentativaDeEnvio = 0;
  
  var limiteDoIndiceParaInicioDeEnvio = 1;
  var limiteDoIndiceParaFimDeEnvio = matrizCompletaDeDadosDaAbaAtual.length - 1;

  for (var indexadorGeralDaLista = limiteDoIndiceParaFimDeEnvio; indexadorGeralDaLista >= limiteDoIndiceParaInicioDeEnvio; indexadorGeralDaLista--) {
    
    var arrayDeUmaLinhaUnica = matrizCompletaDeDadosDaAbaAtual[indexadorGeralDaLista];
    
    var stringNomeDoClienteNaLinha = "Cliente";
    if (arrayDeUmaLinhaUnica[CONFIG.COL_NOME] !== undefined && arrayDeUmaLinhaUnica[CONFIG.COL_NOME] !== null && arrayDeUmaLinhaUnica[CONFIG.COL_NOME] !== "") {
        stringNomeDoClienteNaLinha = arrayDeUmaLinhaUnica[CONFIG.COL_NOME];
    }

    var stringEmailDoClienteNaLinha = "";
    if (arrayDeUmaLinhaUnica[CONFIG.COL_EMAIL] !== undefined && arrayDeUmaLinhaUnica[CONFIG.COL_EMAIL] !== null && arrayDeUmaLinhaUnica[CONFIG.COL_EMAIL] !== "") {
        stringEmailDoClienteNaLinha = arrayDeUmaLinhaUnica[CONFIG.COL_EMAIL].toString().trim();
    }

    var stringPlacaDoClienteNaLinha = "";
    if (arrayDeUmaLinhaUnica[CONFIG.COL_PLACA] !== undefined && arrayDeUmaLinhaUnica[CONFIG.COL_PLACA] !== null && arrayDeUmaLinhaUnica[CONFIG.COL_PLACA] !== "") {
        stringPlacaDoClienteNaLinha = arrayDeUmaLinhaUnica[CONFIG.COL_PLACA].toString().trim().toUpperCase();
    }

    var stringChassiDoClienteNaLinha = "";
    if (arrayDeUmaLinhaUnica[CONFIG.COL_CHASSI] !== undefined && arrayDeUmaLinhaUnica[CONFIG.COL_CHASSI] !== null && arrayDeUmaLinhaUnica[CONFIG.COL_CHASSI] !== "") {
        stringChassiDoClienteNaLinha = arrayDeUmaLinhaUnica[CONFIG.COL_CHASSI].toString().trim().toUpperCase();
    }

    var stringIdentificadorUnicoDaLinha = "";
    if (stringPlacaDoClienteNaLinha !== "") {
        stringIdentificadorUnicoDaLinha = stringPlacaDoClienteNaLinha;
    } else {
        stringIdentificadorUnicoDaLinha = stringChassiDoClienteNaLinha;
    }
    
    var valorDoStatusNaColunaDeEnvio = "";
    if (arrayDeUmaLinhaUnica[CONFIG.COL_ENVIAR] !== undefined && arrayDeUmaLinhaUnica[CONFIG.COL_ENVIAR] !== null && arrayDeUmaLinhaUnica[CONFIG.COL_ENVIAR] !== "") {
        valorDoStatusNaColunaDeEnvio = arrayDeUmaLinhaUnica[CONFIG.COL_ENVIAR].toString().trim().toLowerCase();
    }

    if (valorDoStatusNaColunaDeEnvio !== "enviar") {
        continue;
    }
    
    var variavelDeControleJaEnviouEmailParaOCliente = false;
    if (arrayDeUmaLinhaUnica[CONFIG.COL_CHECK_EMAIL] === true || arrayDeUmaLinhaUnica[CONFIG.COL_CHECK_EMAIL] === "TRUE") {
        variavelDeControleJaEnviouEmailParaOCliente = true;
    }

    var variavelDeControleClienteComEstadoTravado = false;
    if (arrayDeUmaLinhaUnica[CONFIG.COL_TRAVA_ESTADO] === true || arrayDeUmaLinhaUnica[CONFIG.COL_TRAVA_ESTADO] === "TRUE") {
        variavelDeControleClienteComEstadoTravado = true;
    }

    var variavelDeControleClienteComFipeBaixa = false;
    if (arrayDeUmaLinhaUnica[CONFIG.COL_FIPE_BAIXA] === true || arrayDeUmaLinhaUnica[CONFIG.COL_FIPE_BAIXA] === "TRUE") {
        variavelDeControleClienteComFipeBaixa = true;
    }
    
    if (stringIdentificadorUnicoDaLinha === "" || variavelDeControleJaEnviouEmailParaOCliente === true || variavelDeControleClienteComEstadoTravado === true) {
        continue;
    }
    
    if (stringEmailDoClienteNaLinha === "" || regexPadraoDeFormatoDeEmail.test(stringEmailDoClienteNaLinha) === false) {
      sinalizarErroPlanilha(abaPlanilhaAtualNaSessao, indexadorGeralDaLista + 1, "E-mail inválido", stringDaDataFormatadaParaRegistro);
      totalDeErrosNaTentativaDeEnvio++;
      continue;
    }
    
    try {
      
      var textoInicialParaOAssunto = "";
      if (aAbaAtualEhParaExpirados === true) {
          textoInicialParaOAssunto = "[PRAZO EXPIRADO] ";
      }

      var stringDoAssuntoCompletoDoEmail = textoInicialParaOAssunto + "Verificação do sinal do seu rastreador – Veículo: " + stringIdentificadorUnicoDaLinha;
      
      var stringDoTituloInternoDoHTML = "";
      if (aAbaAtualEhParaExpirados === true) {
          stringDoTituloInternoDoHTML = "Prazo Encerrado - Verificação";
      } else {
          stringDoTituloInternoDoHTML = "Aviso Importante - Verificação de Sinal";
      }

      var stringComOTextoPuro = getTextoVerificacaoSinalEmail(stringNomeDoClienteNaLinha, stringIdentificadorUnicoDaLinha, variavelDeControleClienteComFipeBaixa, aAbaAtualEhParaExpirados);
      var stringComOCorpoFormatadoEmHTML = formatarComoEmail(stringComOTextoPuro, stringDoTituloInternoDoHTML);
      
      var objetoComConfiguracoesAvancadasDeEmail = { 
          htmlBody: stringComOCorpoFormatadoEmHTML 
      };

      GmailApp.sendEmail(stringEmailDoClienteNaLinha, stringDoAssuntoCompletoDoEmail, stringComOTextoPuro, objetoComConfiguracoesAvancadasDeEmail);
      
      var celulaNaPlanilhaAlvoStatusEnvio = abaPlanilhaAtualNaSessao.getRange(indexadorGeralDaLista + 1, CONFIG.COL_ENVIAR + 1);
      celulaNaPlanilhaAlvoStatusEnvio.setValue("Enviado");

      var celulaNaPlanilhaAlvoCheckboxEmail = abaPlanilhaAtualNaSessao.getRange(indexadorGeralDaLista + 1, CONFIG.COL_CHECK_EMAIL + 1);
      celulaNaPlanilhaAlvoCheckboxEmail.setValue(true); 
      
      var celulaNaPlanilhaAlvoDataEmail = abaPlanilhaAtualNaSessao.getRange(indexadorGeralDaLista + 1, CONFIG.COL_DATA_EMAIL + 1);
      celulaNaPlanilhaAlvoDataEmail.setValue(stringDaDataFormatadaParaRegistro);

      var celulaNaPlanilhaAlvoResponsavelEmail = abaPlanilhaAtualNaSessao.getRange(indexadorGeralDaLista + 1, CONFIG.COL_RESPONSAVEL_EMAIL + 1);
      celulaNaPlanilhaAlvoResponsavelEmail.setValue(nomeDoResponsavelRecebido);

      var celulaNaPlanilhaAlvoEnderecoEmailCliente = abaPlanilhaAtualNaSessao.getRange(indexadorGeralDaLista + 1, CONFIG.COL_EMAIL + 1);
      celulaNaPlanilhaAlvoEnderecoEmailCliente.setFontColor("#000000").setFontWeight("normal");
      
      SpreadsheetApp.flush(); 
      
      arrayDeUmaLinhaUnica[CONFIG.COL_ENVIAR] = "Enviado";
      arrayDeUmaLinhaUnica[CONFIG.COL_CHECK_EMAIL] = true;
      arrayDeUmaLinhaUnica[CONFIG.COL_DATA_EMAIL] = stringDaDataFormatadaParaRegistro;
      arrayDeUmaLinhaUnica[CONFIG.COL_RESPONSAVEL_EMAIL] = nomeDoResponsavelRecebido;
      
      var stringComONomeAcaoParaAuditoria = "";
      if (aAbaAtualEhParaExpirados === true) {
          stringComONomeAcaoParaAuditoria = "E-mail Enviado (Expirado)";
      } else {
          stringComONomeAcaoParaAuditoria = "E-mail Enviado (Verificação)";
      }

      registrarAuditoriaLog(arrayDeUmaLinhaUnica, stringComONomeAcaoParaAuditoria, stringDaDataFormatadaParaRegistro, nomeDoResponsavelRecebido);
      
      totalDeEnviosComSucesso++;
      Utilities.sleep(2000); 
      
    } catch (erroNaTentativaDoDisparo) {
      sinalizarErroPlanilha(abaPlanilhaAtualNaSessao, indexadorGeralDaLista + 1, "Falha de Envio: " + erroNaTentativaDoDisparo.message, stringDaDataFormatadaParaRegistro);
      totalDeErrosNaTentativaDeEnvio++;
    }
  }
  
  executarGeradorLinksManual();
  
  Browser.msgBox("✅ Relatório (" + nomeDoResponsavelRecebido + "):\nEnviados: " + totalDeEnviosComSucesso + "\nErros: " + totalDeErrosNaTentativaDeEnvio);
}

function formatarComoEmail(stringTextoHtmlOriginal, stringDoTituloDoEmail) {
  var stringTextoComQuebraDeLinhaAjustada = stringTextoHtmlOriginal.replace(/\n/g, '<br>');
  var numeroDoIdentificadorAntiSpam = new Date().getTime();
  
  var codigoHTMLConstruidoFinal = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; max-width: 600px; margin: 0; line-height: 1.6;">
      <h3 style="color: #333333; margin-bottom: 20px; font-weight: bold; text-transform: uppercase;">${stringDoTituloDoEmail}</h3>
      <div style="margin-bottom: 20px;">${stringTextoComQuebraDeLinhaAjustada}</div>
      <div style="border-top: 1px solid #dddddd; padding-top: 15px; margin-top: 20px; font-size: 13px; color: #666666;">
        <img src="https://www.zensegurosbr.com/uploads/images/configuracoes/redimencionar-230-78-logo.png" width="160" alt="ZEN Seguros" style="display: block; margin-bottom: 8px; border: none;">
        Atenciosamente,<br><strong style="color: #444444; font-size: 14px;">Setor de Rastreamento</strong><br>ZEN Seguros<br>(21) 3583-6320 | (21) 97222-0381
      </div>
      <div style="display:none; color:transparent; font-size:1px;">ID de Controle de Qualidade de Servidor: ${numeroDoIdentificadorAntiSpam}</div>
    </div>`;

  return codigoHTMLConstruidoFinal;
}

function getTextoVerificacaoSinalEmail(stringDoNomeCliente, stringDoIdentificadorDoVeiculo, variavelIsFipeBaixaConfirmada, variavelIsAbaDeExpiradoConfirmada) {
  
  var textoSaudacaoFormatado = `Olá, ${stringDoNomeCliente}!\n\nTudo bem?\n\n`;
  var textoDeIntroducaoFormatado = "";

  if (variavelIsAbaDeExpiradoConfirmada === true) {
    textoDeIntroducaoFormatado = `Verificamos que o prazo de 7 dias úteis para a validação/manutenção do rastreador do seu veículo (${stringDoIdentificadorDoVeiculo}) expirou.\n\nAinda não identificamos a normalização do sinal em nossos sistemas.\n\nPara nos ajudar, poderia nos informar por aqui:\nQual a data e o horário aproximado da última vez que o veículo esteve em circulação?\n\n`;
  } else {
    textoDeIntroducaoFormatado = `Estamos entrando em contato para realizar uma verificação de rotina no sinal do rastreador instalado em seu veículo (${stringDoIdentificadorDoVeiculo}).\n\nPara nos ajudar na validação, poderia nos informar por aqui:\nQual a data e o horário aproximado da última vez que o veículo esteve em circulação?\n\n`;
  }
  
  var textoDeAlertaDaCobertura = "";
  if (variavelIsAbaDeExpiradoConfirmada === true) {
    textoDeAlertaDaCobertura = `Prazo Expirado:\nConforme informado anteriormente, a cobertura para sinistros de roubo e furto encontra-se suspensa até a regularização. Por favor, entre em contato com urgência.\n\n`;
  } else {
    textoDeAlertaDaCobertura = `Prazo para Manutenção:\nInformamos que, caso seja constatada a necessidade de intervenção técnica, haverá um prazo de 7 dias úteis. Após este período, a cobertura poderá ser suspensa.\n\n`;
  }
  
  var textoComOsLinksDosAplicativos = `Acesso ao Monitoramento:\nApp Zen Seguros (Android/iOS) ou App Rede Veículos. Fico à disposição para auxiliar!`;
  
  var mensagemComposicaoFinal = "";

  if (variavelIsFipeBaixaConfirmada === true) {
    mensagemComposicaoFinal = textoSaudacaoFormatado + textoDeIntroducaoFormatado + textoComOsLinksDosAplicativos;
  } else {
    mensagemComposicaoFinal = textoSaudacaoFormatado + textoDeIntroducaoFormatado + textoDeAlertaDaCobertura + textoComOsLinksDosAplicativos;
  }

  return mensagemComposicaoFinal;
}

function executarGeradorLinksMenu() {
  executarGeradorLinksManual();
  Browser.msgBox("✅ Links dinâmicos do WhatsApp recriados!");
}

function executarGeradorLinksManual() {
  var abaQueEstaAtivaNaTelaParaOsLinks = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var matrizDosDadosAtuaisDaAbaParaOsLinks = abaQueEstaAtivaNaTelaParaOsLinks.getDataRange().getValues();
  
  var variavelSeAbaExpirado = false;
  if (abaQueEstaAtivaNaTelaParaOsLinks.getName().toLowerCase().indexOf("expirado") > -1) {
      variavelSeAbaExpirado = true;
  }
  
  var numeroLimiteInicialParaOLoop = 1;
  var numeroLimiteFinalParaOLoop = matrizDosDadosAtuaisDaAbaParaOsLinks.length - 1;

  for (var varredorDeLinhasDoWhatsApp = numeroLimiteInicialParaOLoop; varredorDeLinhasDoWhatsApp <= numeroLimiteFinalParaOLoop; varredorDeLinhasDoWhatsApp++) {
    
    var stringDoNumeroDeTelefoneSujo = matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_TELEFONE];
    
    if (stringDoNumeroDeTelefoneSujo !== undefined && stringDoNumeroDeTelefoneSujo !== null && stringDoNumeroDeTelefoneSujo !== "") {
      
      var stringDoNumeroTotalmenteLimpo = stringDoNumeroDeTelefoneSujo.toString().replace(/\D/g, "");
      
      if (stringDoNumeroTotalmenteLimpo.length >= 10) {
        
        if (stringDoNumeroTotalmenteLimpo.startsWith("55") === false) {
            stringDoNumeroTotalmenteLimpo = "55" + stringDoNumeroTotalmenteLimpo;
        }
        
        var stringDaPlacaTemporaria = "";
        if (matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_PLACA] !== undefined && matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_PLACA] !== null) {
            stringDaPlacaTemporaria = matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_PLACA].toString().trim().toUpperCase();
        }

        var stringDoChassiTemporario = "";
        if (matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_CHASSI] !== undefined && matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_CHASSI] !== null) {
            stringDoChassiTemporario = matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_CHASSI].toString().trim().toUpperCase();
        }

        var stringDoIdentificadorGeralDoVeiculo = "";
        if (stringDaPlacaTemporaria !== "") {
            stringDoIdentificadorGeralDoVeiculo = stringDaPlacaTemporaria;
        } else {
            stringDoIdentificadorGeralDoVeiculo = stringDoChassiTemporario;
        }

        var variavelIndicadorSeAbaTemFipeBaixa = false;
        if (matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_FIPE_BAIXA] === true || matrizDosDadosAtuaisDaAbaParaOsLinks[varredorDeLinhasDoWhatsApp][CONFIG.COL_FIPE_BAIXA] === "TRUE") {
            variavelIndicadorSeAbaTemFipeBaixa = true;
        }
        
        var stringDoTextoBrutoParaCodificar = "";
        if (variavelIndicadorSeAbaTemFipeBaixa === true) {
            stringDoTextoBrutoParaCodificar = `Olá! Verificação de rotina do rastreador (${stringDoIdentificadorGeralDoVeiculo}). Qual a última data/hora de circulação?`;
        } else {
            stringDoTextoBrutoParaCodificar = `Olá! Verificação de rotina do rastreador (${stringDoIdentificadorGeralDoVeiculo}). Qual a última data/hora de circulação? (Aviso: Risco de suspensão de cobertura em 7 dias).`;
        }

        var stringDaMensagemCodificadaParaAUrl = encodeURIComponent(stringDoTextoBrutoParaCodificar);
        var stringDoFormatoLivreDeHyperlink = stringDoNumeroDeTelefoneSujo.toString().replace(/=HYPERLINK\(".*?";\s*"(.*?)"\)/, "$1");
        
        var stringDaUrlFinalPersonalizada = '=HYPERLINK("https://wa.me/' + stringDoNumeroTotalmenteLimpo + '?text=' + stringDaMensagemCodificadaParaAUrl + '"; "' + stringDoFormatoLivreDeHyperlink + '")';

        var celulaDoTelefoneParaInjetarNaPlanilha = abaQueEstaAtivaNaTelaParaOsLinks.getRange(varredorDeLinhasDoWhatsApp + 1, CONFIG.COL_TELEFONE + 1);
        celulaDoTelefoneParaInjetarNaPlanilha.setFormula(stringDaUrlFinalPersonalizada);
      }
    }
  }
}

// =========================================================================
// 🧹 BLOCO 6: LIMPEZA, ORGANIZAÇÃO & REFATORAÇÕES NATIVAS
// =========================================================================

function classificarEnviosPorStatusSGA() {
  var abaDoUsuarioPrincipalVisivel = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numeroDaUltimaLinhaExistenteNaAba = abaDoUsuarioPrincipalVisivel.getLastRow();
  
  if (numeroDaUltimaLinhaExistenteNaAba <= CONFIG.LINHAS_CABECHALHO) {
      return;
  }

  var regiaoDaAbaInteriraAcessada = abaDoUsuarioPrincipalVisivel.getRange(1, 1, numeroDaUltimaLinhaExistenteNaAba, CONFIG.QTD_COLUNAS);
  var matrizDeDadosNativosRetornados = regiaoDaAbaInteriraAcessada.getValues();
  
  var variavelSeHouveMudancas = false;
  var somatorioTotalDeSucessosNaClassificacao = 0;

  for (var varredorDeLinhaDeClassificacao = CONFIG.LINHAS_CABECHALHO; varredorDeLinhaDeClassificacao < matrizDeDadosNativosRetornados.length; varredorDeLinhaDeClassificacao++) {
    
    var stringDoStatusDeEnvioAtual = "";
    if (matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_ENVIAR] !== undefined && matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_ENVIAR] !== null) {
        stringDoStatusDeEnvioAtual = matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_ENVIAR].toString().trim().toLowerCase();
    }

    var stringDaClassificacaoAtualDoVeiculo = "";
    if (matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_CLASSIFICACAO] !== undefined && matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_CLASSIFICACAO] !== null) {
        stringDaClassificacaoAtualDoVeiculo = matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_CLASSIFICACAO].toString().trim().toLowerCase();
    }

    var stringDaSituacaoAtualDoVeiculo = "";
    if (matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_SITUACAO] !== undefined && matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_SITUACAO] !== null) {
        stringDaSituacaoAtualDoVeiculo = matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_SITUACAO].toString().trim().toUpperCase();
    }

    if (stringDoStatusDeEnvioAtual === "enviar" || stringDoStatusDeEnvioAtual === "não enviar" || stringDoStatusDeEnvioAtual === "nao enviar" || stringDoStatusDeEnvioAtual === "enviado") {
        continue;
    }

    if (stringDaClassificacaoAtualDoVeiculo !== "concluído" && stringDaClassificacaoAtualDoVeiculo !== "concluido") {
       matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_ENVIAR] = "Verificar"; 
       variavelSeHouveMudancas = true;
       somatorioTotalDeSucessosNaClassificacao++;
       continue;
    }

    if (stringDaSituacaoAtualDoVeiculo === "ATIVO") {
      matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_ENVIAR] = "enviar"; 
    } else {
      matrizDeDadosNativosRetornados[varredorDeLinhaDeClassificacao][CONFIG.COL_ENVIAR] = "Não enviar"; 
    }
    
    variavelSeHouveMudancas = true;
    somatorioTotalDeSucessosNaClassificacao++;
  }

  if (variavelSeHouveMudancas === true) {
    regiaoDaAbaInteriraAcessada.setValues(matrizDeDadosNativosRetornados);
    SpreadsheetApp.flush();
    Browser.msgBox("🚦 " + somatorioTotalDeSucessosNaClassificacao + " linhas classificadas com sucesso.");
  }
}

function marcarFipeBaixaNativo() {
  var abaMestreDaTelaAtual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var blocoCompletoDaMatrizValores = abaMestreDaTelaAtual.getDataRange().getValues();
  
  var numeroDelimitadorStart = CONFIG.LINHAS_CABECHALHO;
  var numeroDelimitadorEnd = blocoCompletoDaMatrizValores.length - 1;
  var numeroDeVeiculosCatalogadosComFipeBaixa = 0;

  for (var w = numeroDelimitadorStart; w <= numeroDelimitadorEnd; w++) {
    
    var stringDaFipeExtraida = "";
    if (blocoCompletoDaMatrizValores[w][CONFIG.COL_FIPE] !== undefined && blocoCompletoDaMatrizValores[w][CONFIG.COL_FIPE] !== null) {
        stringDaFipeExtraida = blocoCompletoDaMatrizValores[w][CONFIG.COL_FIPE].toString().trim();
    }

    if (stringDaFipeExtraida === "") {
        continue;
    }

    var fipeStringProntaParaCalculo = stringDaFipeExtraida.replace(/[^\d.,]/g, "").replace(/\./g, "").replace(",", ".");
    var valorFipeMatematicoCalculado = parseFloat(fipeStringProntaParaCalculo);
    
    if (isNaN(valorFipeMatematicoCalculado) === true) {
        continue;
    }

    var confirmacaoSeVeiculoEhMotocicleta = false;
    if (blocoCompletoDaMatrizValores[w][CONFIG.COL_CLASSIFICACAO] !== undefined && blocoCompletoDaMatrizValores[w][CONFIG.COL_CLASSIFICACAO] !== null) {
        if (blocoCompletoDaMatrizValores[w][CONFIG.COL_CLASSIFICACAO].toString().toUpperCase().indexOf("MOTO") > -1) {
            confirmacaoSeVeiculoEhMotocicleta = true;
        }
    }

    var numeroDoTetoMaximoSeguro = 30000;
    if (confirmacaoSeVeiculoEhMotocicleta === true) {
        numeroDoTetoMaximoSeguro = 20000;
    }

    if (valorFipeMatematicoCalculado < numeroDoTetoMaximoSeguro) {
      var celulaNaColunaDaFipeBaixa = abaMestreDaTelaAtual.getRange(w + 1, CONFIG.COL_FIPE_BAIXA + 1);
      celulaNaColunaDaFipeBaixa.setValue(true);
      numeroDeVeiculosCatalogadosComFipeBaixa++;
    }
  }
  
  SpreadsheetApp.flush();
  Browser.msgBox("✅ " + numeroDeVeiculosCatalogadosComFipeBaixa + " veículos marcados como FIPE Baixa (Leitura Nativa).");
}

function verificarEstadoNativo() {
  var telaPlanilhaAtual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var arrayTodosOsDadosDaTela = telaPlanilhaAtual.getDataRange().getValues();
  
  var numeroIdxComeco = CONFIG.LINHAS_CABECHALHO;
  var numeroIdxTermino = arrayTodosOsDadosDaTela.length - 1;
  var somatorioDeCarrosBloqueadosForaDaSede = 0;

  for (var indexadorDeEstados = numeroIdxComeco; indexadorDeEstados <= numeroIdxTermino; indexadorDeEstados++) {
    
    var stringDoEstadoEncontradoNoLoop = "";
    if (arrayTodosOsDadosDaTela[indexadorDeEstados][CONFIG.COL_ESTADO] !== undefined && arrayTodosOsDadosDaTela[indexadorDeEstados][CONFIG.COL_ESTADO] !== null) {
        stringDoEstadoEncontradoNoLoop = arrayTodosOsDadosDaTela[indexadorDeEstados][CONFIG.COL_ESTADO].toString().trim().toUpperCase();
    }

    if (stringDoEstadoEncontradoNoLoop !== "") {
        if (stringDoEstadoEncontradoNoLoop !== "RJ" && stringDoEstadoEncontradoNoLoop !== "N/A" && stringDoEstadoEncontradoNoLoop !== "RIO DE JANEIRO") {
            
            var celulaNaColunaDaTravaDeEstado = telaPlanilhaAtual.getRange(indexadorDeEstados + 1, CONFIG.COL_TRAVA_ESTADO + 1);
            celulaNaColunaDaTravaDeEstado.setValue(true);
            
            somatorioDeCarrosBloqueadosForaDaSede++;
        }
    }
  }
  
  SpreadsheetApp.flush();
  Browser.msgBox("📍 " + somatorioDeCarrosBloqueadosForaDaSede + " clientes fora do RJ travados.");
}

function sincronizarComentariosOcultos() {
  var abaLocalSincronizadora = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numeroDaUltimaLinhaPossivelDaAba = abaLocalSincronizadora.getLastRow();
  
  if (numeroDaUltimaLinhaPossivelDaAba < 2) {
      return;
  }

  var zonaDePesquisaGeral = abaLocalSincronizadora.getRange(1, 1, numeroDaUltimaLinhaPossivelDaAba, 27);
  var matrizDosDadosNusEMCrus = zonaDePesquisaGeral.getValues();
  
  var bibliotecaNomesDasColunasOcultasParaAsNotas = { 
      1: "Entrada", 
      5: "Fipe", 
      6: "Email", 
      8: "Bairro", 
      9: "Cidade", 
      10: "Estado", 
      12: "Situação", 
      15: "GD2", 
      16: "GD3", 
      20: "Data EM", 
      21: "Data WPP", 
      22: "Resp EM", 
      23: "Resp WPP", 
      25: "Trava", 
      26: "FipeBx" 
  };

  var arrayDaSequenciaColunasOcultas = [1, 5, 6, 8, 9, 10, 12, 15, 16, 20, 21, 22, 23, 25, 26];
  
  var agrupamentoGeralDaMatrizDeNotas = [[""]]; 

  for (var varreduraDasNotas = 1; varreduraDasNotas < matrizDosDadosNusEMCrus.length; varreduraDasNotas++) {
    
    var stringAcumuladoraDoTextoFinalDaNota = "";
    var variavelAchouAlgumaInformacaoDeValor = false;

    for (var ponteiroDeColunaDaNota = 0; ponteiroDeColunaDaNota < arrayDaSequenciaColunasOcultas.length; ponteiroDeColunaDaNota++) {
      
      var numeroDoIndexMatematicoDaColuna = arrayDaSequenciaColunasOcultas[ponteiroDeColunaDaNota];
      var valorCapturadoNaCelulaFocada = matrizDosDadosNusEMCrus[varreduraDasNotas][numeroDoIndexMatematicoDaColuna];
      
      if (valorCapturadoNaCelulaFocada !== "" && valorCapturadoNaCelulaFocada !== null && valorCapturadoNaCelulaFocada !== undefined) {
        
        if (valorCapturadoNaCelulaFocada instanceof Date) {
            valorCapturadoNaCelulaFocada = Utilities.formatDate(valorCapturadoNaCelulaFocada, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (typeof valorCapturadoNaCelulaFocada === "boolean") {
            if (valorCapturadoNaCelulaFocada === true) {
                valorCapturadoNaCelulaFocada = "Sim";
            } else {
                valorCapturadoNaCelulaFocada = "Não";
            }
        }
        
        stringAcumuladoraDoTextoFinalDaNota = stringAcumuladoraDoTextoFinalDaNota + "📌 " + bibliotecaNomesDasColunasOcultasParaAsNotas[numeroDoIndexMatematicoDaColuna] + ": " + valorCapturadoNaCelulaFocada + "\n";
        
        variavelAchouAlgumaInformacaoDeValor = true;
      }
    }
    
    if (variavelAchouAlgumaInformacaoDeValor === true) {
        agrupamentoGeralDaMatrizDeNotas.push([stringAcumuladoraDoTextoFinalDaNota.trim()]);
    } else {
        agrupamentoGeralDaMatrizDeNotas.push([""]);
    }
  }

  var regiaoExataDasPlacasParaInjetarAsNotas = abaLocalSincronizadora.getRange(1, 4, numeroDaUltimaLinhaPossivelDaAba, 1);
  regiaoExataDasPlacasParaInjetarAsNotas.clearNote();
  regiaoExataDasPlacasParaInjetarAsNotas.setNotes(agrupamentoGeralDaMatrizDeNotas);
  
  Browser.msgBox("✅ Comentários sincronizados na Coluna D (Placa).");
}

function atualizarVerificacaoDeDuplicados() {
  var objetoDaControleAbaAtual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (objetoDaControleAbaAtual.getLastRow() < 2) {
      return;
  }
  
  var objetoEspacoParaAAnalise = objetoDaControleAbaAtual.getRange(2, 1, objetoDaControleAbaAtual.getLastRow() - 1, objetoDaControleAbaAtual.getLastColumn());
  var matrizComAInformacaoExtraida = objetoEspacoParaAAnalise.getValues();
  
  var dicionarioDeContagemRegistrosDaPlaca = {}; 
  var dicionarioDeContagemRegistrosDoChassi = {};
  var arrayDeResultadosEmpacotadosParaEnvio = [];
  
  for (var w = 0; w < matrizComAInformacaoExtraida.length; w++) {
    
    var stringDePlacaDetectada = "";
    if (matrizComAInformacaoExtraida[w][CONFIG.COL_PLACA] !== undefined && matrizComAInformacaoExtraida[w][CONFIG.COL_PLACA] !== null) {
        stringDePlacaDetectada = matrizComAInformacaoExtraida[w][CONFIG.COL_PLACA].toString().trim().toUpperCase();
    }

    var stringDeChassiDetectado = "";
    if (matrizComAInformacaoExtraida[w][CONFIG.COL_CHASSI] !== undefined && matrizComAInformacaoExtraida[w][CONFIG.COL_CHASSI] !== null) {
        stringDeChassiDetectado = matrizComAInformacaoExtraida[w][CONFIG.COL_CHASSI].toString().trim().toUpperCase();
    }
    
    if (stringDePlacaDetectada !== "") {
        if (dicionarioDeContagemRegistrosDaPlaca[stringDePlacaDetectada] === undefined) {
            dicionarioDeContagemRegistrosDaPlaca[stringDePlacaDetectada] = 0;
        }
        dicionarioDeContagemRegistrosDaPlaca[stringDePlacaDetectada] = dicionarioDeContagemRegistrosDaPlaca[stringDePlacaDetectada] + 1;
    }

    if (stringDeChassiDetectado !== "") {
        if (dicionarioDeContagemRegistrosDoChassi[stringDeChassiDetectado] === undefined) {
            dicionarioDeContagemRegistrosDoChassi[stringDeChassiDetectado] = 0;
        }
        dicionarioDeContagemRegistrosDoChassi[stringDeChassiDetectado] = dicionarioDeContagemRegistrosDoChassi[stringDeChassiDetectado] + 1;
    }
  }
  
  for (var y = 0; y < matrizComAInformacaoExtraida.length; y++) {
    
    var stringDaNovaPlacaLida = "";
    if (matrizComAInformacaoExtraida[y][CONFIG.COL_PLACA] !== undefined && matrizComAInformacaoExtraida[y][CONFIG.COL_PLACA] !== null) {
        stringDaNovaPlacaLida = matrizComAInformacaoExtraida[y][CONFIG.COL_PLACA].toString().trim().toUpperCase();
    }

    var stringDoNovoChassiLido = "";
    if (matrizComAInformacaoExtraida[y][CONFIG.COL_CHASSI] !== undefined && matrizComAInformacaoExtraida[y][CONFIG.COL_CHASSI] !== null) {
        stringDoNovoChassiLido = matrizComAInformacaoExtraida[y][CONFIG.COL_CHASSI].toString().trim().toUpperCase();
    }
    
    if (stringDaNovaPlacaLida === "" && stringDoNovoChassiLido === "") {
        arrayDeResultadosEmpacotadosParaEnvio.push([""]);
    } else {
        var variavelAPlacaSeRepete = false;
        if (stringDaNovaPlacaLida !== "" && dicionarioDeContagemRegistrosDaPlaca[stringDaNovaPlacaLida] > 1) {
            variavelAPlacaSeRepete = true;
        }

        var variavelOChassiSeRepete = false;
        if (stringDoNovoChassiLido !== "" && dicionarioDeContagemRegistrosDoChassi[stringDoNovoChassiLido] > 1) {
            variavelOChassiSeRepete = true;
        }

        if (variavelAPlacaSeRepete === true || variavelOChassiSeRepete === true) {
            arrayDeResultadosEmpacotadosParaEnvio.push(["Repetido"]);
        } else {
            arrayDeResultadosEmpacotadosParaEnvio.push(["Único"]);
        }
    }
  }
  
  var intervaloDaColunaDestinoDasMarcacoes = objetoDaControleAbaAtual.getRange(2, 1, arrayDeResultadosEmpacotadosParaEnvio.length, 1);
  intervaloDaColunaDestinoDasMarcacoes.setValues(arrayDeResultadosEmpacotadosParaEnvio);
  
  Browser.msgBox("✅ Verificação concluída na Coluna A.");
}

function removerDuplicadosPlacaChassi() {
  var interfaceGrafica = SpreadsheetApp.getUi();
  var confirmacaoUsuario = interfaceGrafica.alert("🗑️ Excluir Duplicadas", "Deseja excluir linhas repetidas mantendo a do topo?", interfaceGrafica.ButtonSet.YES_NO);
  
  if (confirmacaoUsuario !== interfaceGrafica.Button.YES) {
      return;
  }

  var abaMestreAtual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var blocoDados = abaMestreAtual.getRange(2, 1, abaMestreAtual.getLastRow() - 1, abaMestreAtual.getLastColumn());
  var matrizMestreDados = blocoDados.getValues();
  
  var registroDeElementosVistos = {}; 
  var listaDeLinhasMarcadasParaDelecao = [];
  
  for (var varreduraArray = 0; varreduraArray < matrizMestreDados.length; varreduraArray++) {
    
    var valorDaPlacaParaComparar = "";
    if (matrizMestreDados[varreduraArray][CONFIG.COL_PLACA] !== undefined && matrizMestreDados[varreduraArray][CONFIG.COL_PLACA] !== null) {
        valorDaPlacaParaComparar = matrizMestreDados[varreduraArray][CONFIG.COL_PLACA].toString().trim().toUpperCase();
    }

    var valorDoChassiParaComparar = "";
    if (matrizMestreDados[varreduraArray][CONFIG.COL_CHASSI] !== undefined && matrizMestreDados[varreduraArray][CONFIG.COL_CHASSI] !== null) {
        valorDoChassiParaComparar = matrizMestreDados[varreduraArray][CONFIG.COL_CHASSI].toString().trim().toUpperCase();
    }
    
    var placaJaFoiVistaAntes = false;
    if (valorDaPlacaParaComparar !== "" && registroDeElementosVistos[valorDaPlacaParaComparar] === true) {
        placaJaFoiVistaAntes = true;
    }

    var chassiJaFoiVistoAntes = false;
    if (valorDoChassiParaComparar !== "" && registroDeElementosVistos[valorDoChassiParaComparar] === true) {
        chassiJaFoiVistoAntes = true;
    }

    if (placaJaFoiVistaAntes === true || chassiJaFoiVistoAntes === true) {
        listaDeLinhasMarcadasParaDelecao.push(varreduraArray + 2);
    } else { 
        if (valorDaPlacaParaComparar !== "") {
            registroDeElementosVistos[valorDaPlacaParaComparar] = true; 
        }
        if (valorDoChassiParaComparar !== "") {
            registroDeElementosVistos[valorDoChassiParaComparar] = true; 
        }
    }
  }
  
  for (var contadorDelete = listaDeLinhasMarcadasParaDelecao.length - 1; contadorDelete >= 0; contadorDelete--) {
      abaMestreAtual.deleteRow(listaDeLinhasMarcadasParaDelecao[contadorDelete]);
  }
  
  Browser.msgBox("✅ " + listaDeLinhasMarcadasParaDelecao.length + " linhas excluídas.");
}

// =========================================================================
// 🚚 BLOCO 7: PIPELINE DE CRM E MIGRAÇÕES
// =========================================================================

function migrarConcluidos() {
  var controleSistemasDePlanilhas = SpreadsheetApp.getActiveSpreadsheet();
  var abaDoUsuarioPrincipal = controleSistemasDePlanilhas.getSheetByName(CONFIG.NOME_ABA_PRINCIPAL);
  var abaAlvoDestino = garantirAbaExiste(CONFIG.ABA_CONCLUIDOS, abaDoUsuarioPrincipal);
  
  if (abaDoUsuarioPrincipal === null) {
      return;
  }

  var matrizGeralDeLeitura = abaDoUsuarioPrincipal.getDataRange().getValues();
  var somatorioDosClientesMovidos = 0;

  for (var contagemRegressiva = matrizGeralDeLeitura.length - 1; contagemRegressiva >= 1; contagemRegressiva--) {
    
    var statusMarcadoNaColunaDeEnvio = "";
    if (matrizGeralDeLeitura[contagemRegressiva][CONFIG.COL_ENVIAR] !== undefined && matrizGeralDeLeitura[contagemRegressiva][CONFIG.COL_ENVIAR] !== null) {
        statusMarcadoNaColunaDeEnvio = matrizGeralDeLeitura[contagemRegressiva][CONFIG.COL_ENVIAR].toString().trim().toLowerCase();
    }
    
    if (statusMarcadoNaColunaDeEnvio === "enviado") {
      abaAlvoDestino.appendRow(matrizGeralDeLeitura[contagemRegressiva]);
      abaDoUsuarioPrincipal.deleteRow(contagemRegressiva + 1);
      somatorioDosClientesMovidos++;
    }
  }
  
  Browser.msgBox("🚚 " + somatorioDosClientesMovidos + " clientes movidos para a aba Concluídos.");
}

function executarMigracaoSemRetorno() {
  var arquivoAtivoMestre = SpreadsheetApp.getActiveSpreadsheet();
  
  var abaModeloBaseadoNaPrincipal = arquivoAtivoMestre.getSheetByName(CONFIG.NOME_ABA_PRINCIPAL);
  var abaFinalParaSemRetorno = garantirAbaExiste(CONFIG.ABA_SEM_RETORNO, abaModeloBaseadoNaPrincipal);
  
  var listaDeAbasParaFazerAVarreduraDupla = [CONFIG.NOME_ABA_PRINCIPAL, CONFIG.ABA_CONCLUIDOS];
  var dataRelogioDoSistemaAtualmente = new Date();
  var somatorioTotalMigradosParaSemRetorno = 0;

  for (var a = 0; a < listaDeAbasParaFazerAVarreduraDupla.length; a++) {
    
    var abaTemporariaFocada = arquivoAtivoMestre.getSheetByName(listaDeAbasParaFazerAVarreduraDupla[a]);
    
    if (abaTemporariaFocada === null) {
        continue;
    }

    var matrizTodosDadosNestaAbaAtual = abaTemporariaFocada.getDataRange().getValues();
    
    for (var z = matrizTodosDadosNestaAbaAtual.length - 1; z >= 1; z--) {
      
      var infoDataEmailCliente = matrizTodosDadosNestaAbaAtual[z][CONFIG.COL_DATA_EMAIL];
      var infoDataWhatsCliente = matrizTodosDadosNestaAbaAtual[z][CONFIG.COL_DATA_WHATS];
      
      var booleanRespondeuAosContatos = false;
      if (matrizTodosDadosNestaAbaAtual[z][CONFIG.COL_RESPONDEU] === true || matrizTodosDadosNestaAbaAtual[z][CONFIG.COL_RESPONDEU] === "TRUE") {
          booleanRespondeuAosContatos = true;
      }
      
      var clienteTemQualquerDataDeEnvioMarcada = false;
      if (infoDataEmailCliente !== undefined && infoDataEmailCliente !== null && infoDataEmailCliente !== "") {
          clienteTemQualquerDataDeEnvioMarcada = true;
      } else if (infoDataWhatsCliente !== undefined && infoDataWhatsCliente !== null && infoDataWhatsCliente !== "") {
          clienteTemQualquerDataDeEnvioMarcada = true;
      }

      if (clienteTemQualquerDataDeEnvioMarcada === true && booleanRespondeuAosContatos === false) {
        
        var dadoParaEnviarProCalculo = "";
        if (infoDataEmailCliente !== undefined && infoDataEmailCliente !== null && infoDataEmailCliente !== "") {
            dadoParaEnviarProCalculo = infoDataEmailCliente;
        } else {
            dadoParaEnviarProCalculo = infoDataWhatsCliente;
        }

        var montanteDeDiasCalculados = calcularDiasUteis(dadoParaEnviarProCalculo, dataRelogioDoSistemaAtualmente);
        
        if (montanteDeDiasCalculados >= 7) {
          abaFinalParaSemRetorno.appendRow(matrizTodosDadosNestaAbaAtual[z]);
          abaTemporariaFocada.deleteRow(z + 1);
          somatorioTotalMigradosParaSemRetorno++;
        }
      }
    }
  }
  
  Browser.msgBox("🚚 Varredura Dupla: " + somatorioTotalMigradosParaSemRetorno + " clientes sem retorno há mais de 7 dias movidos para aba Sem Retorno.");
}

function garantirAbaExiste(stringNomeAbaDestino, objetoAbaOrigemComoTemplate) {
  var arquivoSheetGeral = SpreadsheetApp.getActiveSpreadsheet();
  var abaVerificadaPeloNome = arquivoSheetGeral.getSheetByName(stringNomeAbaDestino);
  
  if (abaVerificadaPeloNome === null) {
    
    abaVerificadaPeloNome = arquivoSheetGeral.insertSheet(stringNomeAbaDestino);
    
    if (objetoAbaOrigemComoTemplate !== null && objetoAbaOrigemComoTemplate !== undefined) {
      
      var regiaoOrigemCabecalho = objetoAbaOrigemComoTemplate.getRange(1, 1, 1, CONFIG.QTD_COLUNAS);
      
      var matrizValoresDoCabecalho = regiaoOrigemCabecalho.getValues();
      var formatosVisuaisDoCabecalho = regiaoOrigemCabecalho.getTextStyles();
      var corDeFundoDoCabecalho = regiaoOrigemCabecalho.getBackgrounds();
      
      var regiaoDestinoCabecalho = abaVerificadaPeloNome.getRange(1, 1, 1, CONFIG.QTD_COLUNAS);
      
      regiaoDestinoCabecalho.setValues(matrizValoresDoCabecalho);
      regiaoDestinoCabecalho.setTextStyles(formatosVisuaisDoCabecalho);
      regiaoDestinoCabecalho.setBackgrounds(corDeFundoDoCabecalho);
      
      abaVerificadaPeloNome.setFrozenRows(1);
    }
  }
  
  return abaVerificadaPeloNome;
}

// =========================================================================
// 🛠️ BLOCO 8: FUNÇÕES UTILITÁRIAS E PESCADORES BLINDADOS
// =========================================================================

function registrarAuditoriaLog(arrayComALinhaCompletaDosDados, nomeDoEventoDaAcao, dataEHoraComoString, nomeDoResponsavelDaAcao) {
  var espacoDoSpreadsheetAtivo = SpreadsheetApp.getActiveSpreadsheet();
  var templateReferenciaAba = espacoDoSpreadsheetAtivo.getSheetByName(CONFIG.NOME_ABA_PRINCIPAL);
  var abaAlvoDeAuditoria = garantirAbaExiste(CONFIG.ABA_AUDITORIA, templateReferenciaAba);
  
  var conteudoDaCelulaA1 = abaAlvoDeAuditoria.getRange("A1").getValue();
  
  if (conteudoDaCelulaA1 === "") {
    var regiaoDasNovasColunasPersonalizadas = abaAlvoDeAuditoria.getRange("A1:C1");
    regiaoDasNovasColunasPersonalizadas.setValues([["Data/Hora Evento", "Atendente Logado", "Ação Realizada"]]);
    regiaoDasNovasColunasPersonalizadas.setFontWeight("bold");
  }
  
  var arrayMontadoComAsCabecasDeInfo = [dataEHoraComoString, nomeDoResponsavelDaAcao, nomeDoEventoDaAcao];
  var arrayTotalAgrupado = arrayMontadoComAsCabecasDeInfo.concat(arrayComALinhaCompletaDosDados);

  abaAlvoDeAuditoria.appendRow(arrayTotalAgrupado);
}

function sinalizarErroPlanilha(objetoReferenciaParaAba, idDaLinhaAtual, explicacaoDoMotivo, registroDataAtual) {
  var abaReservadaParaErro = garantirAbaExiste(CONFIG.ABA_ERRO, objetoReferenciaParaAba);
  var intervaloLinhaQueCausouOErro = objetoReferenciaParaAba.getRange(idDaLinhaAtual, 1, 1, objetoReferenciaParaAba.getLastColumn());
  var matrizConteudoDaLinhaComErro = intervaloLinhaQueCausouOErro.getValues();
  
  matrizConteudoDaLinhaComErro[0].push("ERRO: " + explicacaoDoMotivo, registroDataAtual);
  
  abaReservadaParaErro.appendRow(matrizConteudoDaLinhaComErro[0]);
}

function calcularDiasUteis(informacaoInicialDaDataComoStringOuObjeto, dataFinalEmFormatoDeObjetoOficial) {
  if (informacaoInicialDaDataComoStringOuObjeto === undefined || informacaoInicialDaDataComoStringOuObjeto === null || informacaoInicialDaDataComoStringOuObjeto === "") {
      return 0;
  }
  
  var dataFormatadaOficialConvertida = null;
  
  if (informacaoInicialDaDataComoStringOuObjeto instanceof Date) {
    dataFormatadaOficialConvertida = new Date(informacaoInicialDaDataComoStringOuObjeto.getTime());
  } else {
    var textoPuroDaStringRecebida = informacaoInicialDaDataComoStringOuObjeto.toString();
    
    if (textoPuroDaStringRecebida.indexOf("/") > -1) {
      // ✅ CORREÇÃO: .split(" ")[0] para isolar só a data antes do horário, e índices corretos
      var partesDaDataSegmentadaPelaBarra = textoPuroDaStringRecebida.split(" ")[0].split("/");
      var diaRecebido = parseInt(partesDaDataSegmentadaPelaBarra[0], 10);
      var mesRecebidoComCorrecaoMatematica = parseInt(partesDaDataSegmentadaPelaBarra[1], 10) - 1;
      var anoRecebido = parseInt(partesDaDataSegmentadaPelaBarra[2], 10);
      
      dataFormatadaOficialConvertida = new Date(anoRecebido, mesRecebidoComCorrecaoMatematica, diaRecebido);
    } else {
      dataFormatadaOficialConvertida = new Date(textoPuroDaStringRecebida);
    }
  }
  
  if (isNaN(dataFormatadaOficialConvertida.getTime()) === true) {
      return 0;
  }
  
  var dataHojeZeradaParaComporBase = new Date(dataFinalEmFormatoDeObjetoOficial.getTime());
  dataHojeZeradaParaComporBase.setHours(0,0,0,0); 
  
  dataFormatadaOficialConvertida.setHours(0,0,0,0);
  
  var contadorAcumuladorDiasCorridos = 0;
  var iteradorDeDiaAtualEmTempoReal = new Date(dataFormatadaOficialConvertida.getTime());
  
  iteradorDeDiaAtualEmTempoReal.setDate(iteradorDeDiaAtualEmTempoReal.getDate() + 1);
  
  while (iteradorDeDiaAtualEmTempoReal <= dataHojeZeradaParaComporBase) {
    var identificarQualODiaDaSemanaAtualizado = iteradorDeDiaAtualEmTempoReal.getDay();
    
    var ehDomingo = false;
    if (identificarQualODiaDaSemanaAtualizado === 0) { ehDomingo = true; }

    var ehSabado = false;
    if (identificarQualODiaDaSemanaAtualizado === 6) { ehSabado = true; }

    var bateuComDataDeFeriadoNacionalDoBrasil = ehFeriadoBrasileiro(iteradorDeDiaAtualEmTempoReal);

    if (ehDomingo === false && ehSabado === false && bateuComDataDeFeriadoNacionalDoBrasil === false) {
        contadorAcumuladorDiasCorridos++;
    }
    
    iteradorDeDiaAtualEmTempoReal.setDate(iteradorDeDiaAtualEmTempoReal.getDate() + 1);
  }
  
  return contadorAcumuladorDiasCorridos;
}

function ehFeriadoBrasileiro(objetoDataParaInspecionarFeriado) {
  var extrairOAnoCompletoDaVariavel = objetoDataParaInspecionarFeriado.getFullYear();
  var extrairOMesDaVariavelComCorrecao = objetoDataParaInspecionarFeriado.getMonth() + 1;
  var extrairODiaDaVariavel = objetoDataParaInspecionarFeriado.getDate();
  
  var dicionarioDeFeriadosConstantes = ["1/1", "21/4", "1/5", "7/9", "12/10", "2/11", "15/11", "25/12"];
  var stringFormadaDoDiaComOMesNoLoopAtual = extrairODiaDaVariavel + "/" + extrairOMesDaVariavelComCorrecao;

  if (dicionarioDeFeriadosConstantes.indexOf(stringFormadaDoDiaComOMesNoLoopAtual) > -1) {
      return true;
  }
  
  var operacaoMatematicaA = extrairOAnoCompletoDaVariavel % 19;
  var operacaoMatematicaB = Math.floor(extrairOAnoCompletoDaVariavel / 100);
  var operacaoMatematicaC = extrairOAnoCompletoDaVariavel % 100;
  var operacaoMatematicaD = Math.floor(operacaoMatematicaB / 4);
  var operacaoMatematicaE = operacaoMatematicaB % 4;
  var operacaoMatematicaF = Math.floor((operacaoMatematicaB + 8) / 25);
  var operacaoMatematicaG = Math.floor((operacaoMatematicaB - operacaoMatematicaF + 1) / 3);
  var operacaoMatematicaH = (19 * operacaoMatematicaA + operacaoMatematicaB - operacaoMatematicaD - operacaoMatematicaG + 15) % 30;
  var operacaoMatematicaI = Math.floor(operacaoMatematicaC / 4);
  var operacaoMatematicaK = operacaoMatematicaC % 4;
  var operacaoMatematicaL = (32 + 2 * operacaoMatematicaE + 2 * operacaoMatematicaI - operacaoMatematicaH - operacaoMatematicaK) % 7;
  var operacaoMatematicaM = Math.floor((operacaoMatematicaA + 11 * operacaoMatematicaH + 22 * operacaoMatematicaL) / 451);
  var mesResolvidoMisterioDaPascoa = Math.floor((operacaoMatematicaH + operacaoMatematicaL - 7 * operacaoMatematicaM + 114) / 31) - 1;
  var diaResolvidoMisterioDaPascoa = ((operacaoMatematicaH + operacaoMatematicaL - 7 * operacaoMatematicaM + 114) % 31) + 1;
  
  var dataCravadaParaRepresentarAPascoa = new Date(extrairOAnoCompletoDaVariavel, mesResolvidoMisterioDaPascoa, diaResolvidoMisterioDaPascoa);
  
  var dataCravadaParaSextaFeiraDaPaixao = new Date(dataCravadaParaRepresentarAPascoa.getTime()); 
  dataCravadaParaSextaFeiraDaPaixao.setDate(dataCravadaParaRepresentarAPascoa.getDate() - 2);
  
  var dataCravadaParaCarnaval = new Date(dataCravadaParaRepresentarAPascoa.getTime()); 
  dataCravadaParaCarnaval.setDate(dataCravadaParaRepresentarAPascoa.getDate() - 47);
  
  var dataCravadaParaCorpusChristi = new Date(dataCravadaParaRepresentarAPascoa.getTime()); 
  dataCravadaParaCorpusChristi.setDate(dataCravadaParaRepresentarAPascoa.getDate() + 60);
  
  if (extrairODiaDaVariavel === dataCravadaParaSextaFeiraDaPaixao.getDate() && extrairOMesDaVariavelComCorrecao === dataCravadaParaSextaFeiraDaPaixao.getMonth() + 1) {
      return true;
  }
  
  if (extrairODiaDaVariavel === dataCravadaParaCarnaval.getDate() && extrairOMesDaVariavelComCorrecao === dataCravadaParaCarnaval.getMonth() + 1) {
      return true;
  }
  
  if (extrairODiaDaVariavel === dataCravadaParaCorpusChristi.getDate() && extrairOMesDaVariavelComCorrecao === dataCravadaParaCorpusChristi.getMonth() + 1) {
      return true;
  }
  
  return false;
}
