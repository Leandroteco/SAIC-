const ABA_DADOS = "dados_cadastro";
const ABA_VINCULOS = "pessoas_vinculadas";
const ABA_USUARIOS = "usuarios_sistema";
const GOOGLE_CLIENT_ID = "929026048656-ef6g930iicha4bdfa4boh55ninluevfa.apps.googleusercontent.com";

const PERFIL_ADMINISTRADOR = "administrador";
const PERFIL_ATENDENTE = "atendente";

const CABECALHOS_DADOS = [
  "id_atendimento",
  "emailCadastro",
  "tipoAtendimento",
  "motivo",
  "re",
  "nome",
  "cpf",
  "telefone",
  "email",
  "dataIngresso",
  "dataNascimento",
  "sexo",
  "opmAtual",
  "situacaoStatus",
  "dataInatividade",
  "estadoCivil",
  "numeroFilhos",
  "cep",
  "rua",
  "bairro",
  "cidade",
  "estado",
  "numero",
  "complemento",
  "observacoes",
  "responsavel",
  "dataCadastro",
  "postoGraduacao"
];

const CABECALHOS_VINCULOS = [
  "id_vinculo",
  "id_atendimento",
  "nome",
  "cpf",
  "tipoVinculo",
  "parentesco",
  "observacoes"
];

const CABECALHOS_USUARIOS = [
  "email",
  "perfil",
  "ativo",
  "nome"
];

function configurarEstruturaPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDados = obterOuCriarAba(ss, ABA_DADOS, CABECALHOS_DADOS);
  const sheetVinculos = obterOuCriarAba(ss, ABA_VINCULOS, CABECALHOS_VINCULOS);
  const sheetUsuarios = obterOuCriarAba(ss, ABA_USUARIOS, CABECALHOS_USUARIOS);

  return {
    sheetDados: sheetDados,
    sheetVinculos: sheetVinculos,
    sheetUsuarios: sheetUsuarios
  };
}

function obterOuCriarAba(ss, nomeAba, cabecalhos) {
  let sheet = ss.getSheetByName(nomeAba);

  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
    sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function obterEmailUsuarioGoogle() {
  return String(Session.getActiveUser().getEmail() || "")
    .toLowerCase()
    .trim();
}

function obterUsuarioSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obterOuCriarAba(ss, ABA_USUARIOS, CABECALHOS_USUARIOS);
  const email = obterEmailUsuarioGoogle();

  if (!email) {
    return {
      autorizado: false,
      email: "",
      perfil: "",
      nome: "",
      mensagem: "Nao foi possivel identificar o email Google do usuario. Verifique a implantacao do Web App."
    };
  }

  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const emailCadastrado = String(dados[i][0] || "").toLowerCase().trim();
    const perfil = String(dados[i][1] || "").toLowerCase().trim();
    const ativo = String(dados[i][2] || "").toLowerCase().trim();
    const nome = String(dados[i][3] || "").trim();

    if (emailCadastrado === email) {
      if (ativo !== "sim") {
        return {
          autorizado: false,
          email: email,
          perfil: perfil,
          nome: nome,
          mensagem: "Usuario cadastrado, mas inativo."
        };
      }

      return {
        autorizado: true,
        email: email,
        perfil: perfil,
        nome: nome,
        mensagem: ""
      };
    }
  }

  return {
    autorizado: false,
    email: email,
    perfil: "",
    nome: "",
    mensagem: "Email nao autorizado: " + email
  };
}

function validarUsuarioAutorizado() {
  const usuario = obterUsuarioSistema();

  if (!usuario.autorizado) {
    throw new Error(usuario.mensagem || "Acesso negado.");
  }

  return usuario;
}

function validarUsuarioAdministrador() {
  const usuario = validarUsuarioAutorizado();

  if (usuario.perfil !== PERFIL_ADMINISTRADOR) {
    throw new Error("Acesso negado. Somente administradores podem acessar este recurso.");
  }

  return usuario;
}

function renderizarAcessoNegado(usuario) {
  const email = escaparHtmlServidor(usuario.email || "email nao identificado");
  const mensagem = escaparHtmlServidor(usuario.mensagem || "Acesso negado.");

  const html =
    "<!DOCTYPE html>" +
    "<html>" +
    "<head>" +
    "<base target='_top'>" +
    "<style>" +
    "body{background:#f4f8fb;color:#334155;font-family:'Segoe UI',Arial,sans-serif;margin:0;padding:30px;}" +
    ".box{background:white;border:1px solid #d8e2ea;border-radius:12px;margin:60px auto;max-width:560px;padding:28px;text-align:center;}" +
    "h1{color:#b42318;margin-top:0;}" +
    "p{line-height:1.5;}" +
    ".email{background:#f1f5f9;border-radius:8px;display:inline-block;margin-top:10px;padding:8px 12px;}" +
    "</style>" +
    "</head>" +
    "<body>" +
    "<div class='box'>" +
    "<h1>Acesso negado</h1>" +
    "<p>" + mensagem + "</p>" +
    "<div class='email'>" + email + "</div>" +
    "</div>" +
    "</body>" +
    "</html>";

  return HtmlService.createHtmlOutput(html).setTitle("SAIC - Acesso negado");
}

function escaparHtmlServidor(valor) {
  return String(valor === null || valor === undefined ? "" : valor)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function doGet(e) {
  const pagina = e && e.parameter && e.parameter.pagina === "relatorios"
    ? "relatorios"
    : "Index";

  return renderizarPagina(pagina, "", null, "");
}

function obterConfigLoginGoogle() {
  return {
    clientId: GOOGLE_CLIENT_ID
  };
}

function doPost(e) {
  const idToken = e && e.parameter ? String(e.parameter.credential || "") : "";
  const destino = e && e.parameter ? String(e.parameter.state || "index") : "index";
  const pagina = destino === "relatorios" ? "relatorios" : "Index";

  try {
    const usuario = validarLoginGoogle(idToken);

    if (pagina === "relatorios" && usuario.perfil !== "administrador") {
      throw new Error("Acesso permitido somente para administradores.");
    }

    return renderizarPagina(pagina, idToken, usuario, "");
  } catch (erro) {
    return renderizarPagina(pagina, "", null, erro.message);
  }
}

function renderizarPagina(nomeArquivo, tokenGoogle, usuario, mensagemLogin) {
  const template = HtmlService.createTemplateFromFile(nomeArquivo);

  template.googleClientId = GOOGLE_CLIENT_ID;
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.tokenGoogleInicial = tokenGoogle || "";
  template.usuarioInicialJson = JSON.stringify(usuario || null);
  template.mensagemLoginInicial = mensagemLogin || "";

  return template
    .evaluate()
    .setTitle("SAIC")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function validarLoginGoogle(idToken) {
  const usuario = validarUsuarioPorToken(idToken);

  return {
    autorizado: true,
    email: usuario.email,
    perfil: usuario.perfil,
    nome: usuario.nome,
    administrador: usuario.perfil === PERFIL_ADMINISTRADOR
  };
}

function validarUsuarioPorToken(idToken) {
  const dadosToken = verificarTokenGoogle(idToken);
  const usuario = obterUsuarioSistemaPorEmail(dadosToken.email);

  if (!usuario.autorizado) {
    throw new Error(usuario.mensagem || "Acesso negado.");
  }

  return usuario;
}

function validarAdministradorPorToken(idToken) {
  const usuario = validarUsuarioPorToken(idToken);

  if (usuario.perfil !== PERFIL_ADMINISTRADOR) {
    throw new Error("Acesso negado. Somente administradores podem acessar este recurso.");
  }

  return usuario;
}

function verificarTokenGoogle(idToken) {
  if (!idToken) {
    throw new Error("Login Google não informado.");
  }

  const resposta = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken),
    {
      method: "get",
      muteHttpExceptions: true
    }
  );

  if (resposta.getResponseCode() !== 200) {
    throw new Error("Login Google inválido ou expirado.");
  }

  const payload = JSON.parse(resposta.getContentText());

  if (payload.aud !== GOOGLE_CLIENT_ID) {
    throw new Error("Token Google não pertence a este sistema.");
  }

  if (payload.iss !== "https://accounts.google.com" && payload.iss !== "accounts.google.com") {
    throw new Error("Emissor do token Google inválido.");
  }

  if (String(payload.email_verified) !== "true") {
    throw new Error("Email Google não verificado.");
  }

  if (!payload.email) {
    throw new Error("Token Google sem email.");
  }

  return {
    email: String(payload.email).toLowerCase().trim(),
    nome: payload.name || "",
    sub: payload.sub || ""
  };
}

function obterUsuarioSistemaPorEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obterOuCriarAba(ss, ABA_USUARIOS, CABECALHOS_USUARIOS);
  const emailNormalizado = String(email || "").toLowerCase().trim();

  if (!emailNormalizado) {
    return {
      autorizado: false,
      email: "",
      perfil: "",
      nome: "",
      mensagem: "Email não identificado."
    };
  }

  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const emailCadastrado = String(dados[i][0] || "").toLowerCase().trim();
    const perfil = String(dados[i][1] || "").toLowerCase().trim();
    const ativo = String(dados[i][2] || "").toLowerCase().trim();
    const nome = String(dados[i][3] || "").trim();

    if (emailCadastrado === emailNormalizado) {
      if (ativo !== "sim") {
        return {
          autorizado: false,
          email: emailNormalizado,
          perfil: perfil,
          nome: nome,
          mensagem: "Usuário cadastrado, mas inativo."
        };
      }

      return {
        autorizado: true,
        email: emailNormalizado,
        perfil: perfil,
        nome: nome,
        mensagem: ""
      };
    }
  }

  return {
    autorizado: false,
    email: emailNormalizado,
    perfil: "",
    nome: "",
    mensagem: "Email não autorizado: " + emailNormalizado
  };
}


function salvarAtendimento(dados, idToken) {
  const usuario = validarUsuarioPorToken(idToken);
  const estrutura = configurarEstruturaPlanilha();
  const sheet = estrutura.sheetDados;
  const sheetVinculos = estrutura.sheetVinculos;

  if (!sheet) throw new Error("A aba dados_cadastro nao foi encontrada.");
  if (!sheetVinculos) throw new Error("A aba pessoas_vinculadas nao foi encontrada.");

  const idAtendimento = gerarNovoId(sheet, 1);
  const dataCadastro = new Date();

  sheet.appendRow([
    idAtendimento,
    usuario.email,
    normalizar(dados.tipoAtendimento),
    normalizar(dados.motivo),
    normalizar(dados.re),
    normalizar(dados.nome),
    formatarCPF(dados.cpf),
    normalizarTelefone(dados.telefone),
    normalizar(dados.email),
    dados.dataIngresso || "",
    dados.dataNascimento || "",
    normalizar(dados.sexo),
    normalizar(dados.opmAtual),
    normalizar(dados.situacaoStatus),
    dados.dataInatividade || "",
    normalizar(dados.estadoCivil),
    dados.numeroFilhos || "",
    dados.cep || "",
    normalizar(dados.rua),
    normalizar(dados.bairro),
    normalizar(dados.cidade),
    normalizar(dados.estado),
    dados.numero || "",
    normalizar(dados.complemento),
    normalizar(dados.observacoes),
    normalizar(dados.responsavel),
    dataCadastro,
    dados.postoGraduacao || ""
  ]);

  if (dados.pessoasVinculadas && dados.pessoasVinculadas.length > 0) {
    dados.pessoasVinculadas.forEach(function(pessoa) {
      if (pessoa.nome || pessoa.cpf) {
        const idVinculo = gerarNovoId(sheetVinculos, 1);

        sheetVinculos.appendRow([
          idVinculo,
          idAtendimento,
          normalizar(pessoa.nome),
          formatarCPF(pessoa.cpf),
          normalizar(pessoa.tipoVinculo),
          normalizar(pessoa.parentesco),
          normalizar(pessoa.observacoes)
        ]);
      }
    });
  }

  return "Atendimento salvo com sucesso!";
}

function gerarNovoId(sheet, coluna) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return 1;

  const valores = sheet.getRange(2, coluna, lastRow - 1, 1).getValues().flat();
  const numeros = valores.filter(v => !isNaN(v) && v !== "").map(Number);

  return numeros.length ? Math.max(...numeros) + 1 : 1;
}

function normalizar(texto) {
  if (!texto) return "";

  return String(texto)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function somenteNumeros(valor) {
  if (!valor) return "";
  return String(valor).replace(/\D/g, "");
}

function formatarCPF(cpf) {
  cpf = somenteNumeros(cpf);

  if (cpf.length !== 11) return cpf;

  return cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
}

function normalizarTelefone(valor) {
  if (!valor) return "";
  return String(valor).replace(/\D/g, "");
}

function montarRegistro(linha) {
  return {
    id: linha[0],
    emailCadastro: linha[1],
    naps: linha[1],
    tipoAtendimento: linha[2],
    motivo: linha[3],
    re: linha[4],
    nome: linha[5],
    cpf: linha[6],
    telefone: linha[7],
    email: linha[8],
    dataIngresso: formatarDataParaInput(linha[9]),
    dataNascimento: formatarDataParaInput(linha[10]),
    sexo: linha[11],
    opmAtual: linha[12],
    situacaoStatus: linha[13],
    dataInatividade: formatarDataParaInput(linha[14]),
    estadoCivil: linha[15],
    numeroFilhos: linha[16],
    cep: linha[17],
    rua: linha[18],
    bairro: linha[19],
    cidade: linha[20],
    estado: linha[21],
    numero: linha[22],
    complemento: linha[23],
    observacoes: linha[24],
    responsavel: linha[25],
    dataCadastro: linha[26],
    postoGraduacao: linha[27] || ""
  };
}

function formatarDataParaInput(valor) {
  const data = obterData(valor);

  if (!data) return "";

  const ano = data.getFullYear();
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  const dia = String(data.getDate()).padStart(2, "0");

  return ano + "-" + mes + "-" + dia;
}

function buscarCadastro(termo, idToken) {
  validarUsuarioPorToken(idToken);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DADOS);

  if (!sheet) {
    return {
      encontrado: false,
      mensagem: "A aba dados_cadastro nao foi encontrada."
    };
  }

  const linhas = sheet.getDataRange().getValues();
  const buscaTexto = normalizar(termo);
  const buscaNumeros = somenteNumeros(termo);

  let resultados = [];
  let chavesEncontradas = {};

  for (let i = linhas.length - 1; i >= 1; i--) {
    const l = linhas[i];

    const registro = {
      re: l[4] || "",
      postoGraduacao: l[27] || "",
      nome: l[5] || "",
      cpf: l[6] || "",
      telefone: l[7] || "",
      email: l[8] || "",
      dataIngresso: converterData(l[9]),
      dataNascimento: converterData(l[10]),
      sexo: l[11] || "",
      opmAtual: l[12] || "",
      situacaoStatus: l[13] || "",
      dataInatividade: converterData(l[14]),
      estadoCivil: l[15] || "",
      numeroFilhos: l[16] || "",
      cep: l[17] || "",
      rua: l[18] || "",
      bairro: l[19] || "",
      cidade: l[20] || "",
      estado: l[21] || "",
      numero: l[22] || "",
      complemento: l[23] || "",
      dataCadastro: formatarDataBrasil(l[26])
    };

    const reTexto = normalizar(registro.re);
    const nomeTexto = normalizar(registro.nome);
    const cpfNumerico = somenteNumeros(registro.cpf);

    const pesquisouCPF = buscaNumeros.length >= 11;
    const pesquisouRE = /^[0-9]{6}-[0-9A]$/i.test(String(termo).trim());
    const pesquisouNome = buscaTexto.length >= 3 && !pesquisouCPF && !pesquisouRE;

    const achouCPF = pesquisouCPF && cpfNumerico === buscaNumeros;
    const achouRE = pesquisouRE && reTexto === buscaTexto;
    const achouNome = pesquisouNome && nomeTexto.includes(buscaTexto);

    if (achouCPF || achouRE || achouNome) {
      const chaveUnica = cpfNumerico || reTexto || nomeTexto;

      if (!chavesEncontradas[chaveUnica]) {
        chavesEncontradas[chaveUnica] = true;
        resultados.push(registro);
      }
    }
  }

  if (resultados.length === 0) {
    return {
      encontrado: false,
      mensagem: "Nenhum cadastro anterior localizado. Preencha novo cadastro."
    };
  }

  if (buscaNumeros.length >= 11 || /^[0-9]{6}-[0-9A]$/i.test(String(termo).trim())) {
    return {
      encontrado: true,
      multiplos: false,
      registro: resultados[0]
    };
  }

  if (resultados.length === 1) {
    return {
      encontrado: true,
      multiplos: false,
      registro: resultados[0]
    };
  }

  return {
    encontrado: true,
    multiplos: true,
    resultados: resultados.slice(0, 10)
  };
}

function converterData(valor) {
  return formatarDataParaInput(valor);
}

function gerarRelatorioGerencial(filtrosOuDataInicial, dataFinal, tiposRelatorio, idToken) {
  validarAdministradorPorToken(idToken);

  const filtros = normalizarFiltrosRelatorio(filtrosOuDataInicial, dataFinal, tiposRelatorio);
  const estrutura = configurarEstruturaPlanilha();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = estrutura.sheetDados;

  const dados = sheet.getDataRange().getValues();
  const vinculosPorAtendimento = carregarVinculosPorAtendimento(ss);
  const registros = [];

  for (let i = 1; i < dados.length; i++) {
    registros.push(montarRegistroRelatorio(dados[i], vinculosPorAtendimento));
  }

  const filtrosPreparados = prepararFiltrosRelatorio(filtros);
  const filtrados = registros.filter(function(registro) {
    return registroPassaFiltrosRelatorio(registro, filtrosPreparados);
  });

  const limiteDetalhes = Number(filtros.limiteDetalhes) || 800;

  return {
    filtrosAplicados: filtros,
    resumo: montarResumoRelatorio(filtrados),
    distribuicoes: montarDistribuicoesRelatorio(filtrados),
    registros: montarDetalhesRelatorio(filtrados, limiteDetalhes),
    totalRegistros: filtrados.length,
    limiteDetalhes: limiteDetalhes,
    opcoes: montarOpcoesRelatorio(registros, vinculosPorAtendimento)
  };
}

function normalizarFiltrosRelatorio(filtrosOuDataInicial, dataFinal, tiposRelatorio) {
  if (filtrosOuDataInicial && typeof filtrosOuDataInicial === "object" && !Array.isArray(filtrosOuDataInicial)) {
    return filtrosOuDataInicial;
  }

  return {
    dataInicial: filtrosOuDataInicial || "",
    dataFinal: dataFinal || "",
    secoes: tiposRelatorio || []
  };
}

function prepararFiltrosRelatorio(filtros) {
  return {
    dataInicial: obterDataInicio(filtros.dataInicial),
    dataFinal: obterDataFim(filtros.dataFinal),
    buscaLivre: normalizar(filtros.buscaLivre),
    tipoAtendimento: normalizar(filtros.tipoAtendimento),
    motivo: normalizar(filtros.motivo),
    naps: normalizar(filtros.naps),
    opmAtual: normalizar(filtros.opmAtual),
    cidade: normalizar(filtros.cidade),
    bairro: normalizar(filtros.bairro),
    estado: normalizar(filtros.estado),
    responsavel: normalizar(filtros.responsavel),
    situacaoStatus: normalizar(filtros.situacaoStatus),
    sexo: normalizar(filtros.sexo),
    estadoCivil: normalizar(filtros.estadoCivil),
    faixaEtaria: normalizar(filtros.faixaEtaria),
    tempoServico: normalizar(filtros.tempoServico),
    vinculos: normalizar(filtros.vinculos),
    parentesco: normalizar(filtros.parentesco)
  };
}

function registroPassaFiltrosRelatorio(registro, filtros) {
  if (filtros.dataInicial || filtros.dataFinal) {
    if (!registro.dataCadastroData) return false;
    if (filtros.dataInicial && registro.dataCadastroData < filtros.dataInicial) return false;
    if (filtros.dataFinal && registro.dataCadastroData > filtros.dataFinal) return false;
  }

  if (filtros.buscaLivre) {
    const textoBusca = normalizar([
      registro.nome,
      registro.postoGraduacao,
      registro.re,
      registro.cpf,
      registro.tipoAtendimento,
      registro.motivo,
      registro.naps,
      registro.emailCadastro,
      registro.opmAtual,
      registro.cidade,
      registro.bairro,
      registro.estado,
      registro.responsavel,
      registro.situacaoStatus,
      registro.cep
    ].join(" "));

    if (!textoBusca.includes(filtros.buscaLivre)) return false;
  }

  if (!campoIgualRelatorio(registro.tipoAtendimento, filtros.tipoAtendimento)) return false;
  if (!campoContemRelatorio(registro.motivo, filtros.motivo)) return false;
  if (!campoIgualRelatorio(registro.naps, filtros.naps)) return false;
  if (!campoIgualRelatorio(registro.opmAtual, filtros.opmAtual)) return false;
  if (!campoIgualRelatorio(registro.cidade, filtros.cidade)) return false;
  if (!campoIgualRelatorio(registro.bairro, filtros.bairro)) return false;
  if (!campoIgualRelatorio(registro.estado, filtros.estado)) return false;
  if (!campoIgualRelatorio(registro.responsavel, filtros.responsavel)) return false;
  if (!campoIgualRelatorio(registro.situacaoStatus, filtros.situacaoStatus)) return false;
  if (!campoIgualRelatorio(registro.sexo, filtros.sexo)) return false;
  if (!campoIgualRelatorio(registro.estadoCivil, filtros.estadoCivil)) return false;
  if (!campoIgualRelatorio(registro.faixaEtaria, filtros.faixaEtaria)) return false;
  if (!campoIgualRelatorio(registro.tempoServico, filtros.tempoServico)) return false;

  if (filtros.vinculos === "com" && registro.quantidadeVinculos === 0) return false;
  if (filtros.vinculos === "sem" && registro.quantidadeVinculos > 0) return false;

  if (filtros.parentesco && !registro.vinculos.some(function(vinculo) {
    return normalizar(vinculo.parentesco) === filtros.parentesco;
  })) {
    return false;
  }

  return true;
}

function campoIgualRelatorio(valor, filtro) {
  if (!filtro) return true;
  return normalizar(valor) === filtro;
}

function campoContemRelatorio(valor, filtro) {
  if (!filtro) return true;
  return normalizar(valor).includes(filtro);
}

function carregarVinculosPorAtendimento(ss) {
  const sheet = ss.getSheetByName(ABA_VINCULOS);
  const mapa = {};

  if (!sheet) return mapa;

  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const idAtendimento = String(linha[1] || "");

    if (!idAtendimento) continue;

    if (!mapa[idAtendimento]) mapa[idAtendimento] = [];

    mapa[idAtendimento].push({
      id: linha[0] || "",
      idAtendimento: idAtendimento,
      nome: linha[2] || "",
      cpf: linha[3] || "",
      tipoVinculo: linha[4] || "",
      parentesco: linha[5] || "",
      observacoes: linha[6] || ""
    });
  }

  return mapa;
}

function montarRegistroRelatorio(linha, vinculosPorAtendimento) {
  const id = linha[0] || "";
  const dataCadastroData = obterData(linha[26]);
  const dataNascimentoData = obterData(linha[10]);
  const dataIngressoData = obterData(linha[9]);
  const idade = calcularAnosAteHoje(dataNascimentoData);
  const tempoServicoAnos = calcularAnosAteHoje(dataIngressoData);
  const vinculos = vinculosPorAtendimento[String(id)] || [];

  return {
    id: id,
    emailCadastro: linha[1] || "",
    naps: linha[1] || "",
    tipoAtendimento: linha[2] || "",
    motivo: linha[3] || "",
    re: linha[4] || "",
    nome: linha[5] || "",
    postoGraduacao: linha[27] || "",
    cpf: linha[6] || "",
    telefone: linha[7] || "",
    email: linha[8] || "",
    dataIngresso: formatarDataBrasil(linha[9]),
    dataNascimento: formatarDataBrasil(linha[10]),
    idade: idade,
    faixaEtaria: obterFaixaEtaria(idade),
    sexo: linha[11] || "",
    opmAtual: linha[12] || "",
    situacaoStatus: linha[13] || "",
    dataInatividade: formatarDataBrasil(linha[14]),
    estadoCivil: linha[15] || "",
    numeroFilhos: linha[16] || "",
    cep: linha[17] || "",
    rua: linha[18] || "",
    bairro: linha[19] || "",
    cidade: linha[20] || "",
    estado: linha[21] || "",
    numero: linha[22] || "",
    complemento: linha[23] || "",
    observacoes: linha[24] || "",
    responsavel: linha[25] || "",
    dataCadastro: formatarDataBrasil(linha[26]),
    dataCadastroIso: formatarDataParaInput(linha[26]),
    dataCadastroData: dataCadastroData,
    dataCadastroTimestamp: dataCadastroData ? dataCadastroData.getTime() : 0,
    mesCadastro: obterMesAno(dataCadastroData),
    tempoServico: obterFaixaTempoServico(tempoServicoAnos),
    endereco: montarEnderecoRelatorio(linha),
    vinculos: vinculos,
    quantidadeVinculos: vinculos.length
  };
}

function montarResumoRelatorio(registros) {
  const pessoas = {};
  const atendimentosPorPessoa = {};
  let totalVinculos = 0;
  let atendimentosComVinculos = 0;
  let atendimentosFamiliaresOuGrupo = 0;

  registros.forEach(function(registro) {
    const chave = obterChavePessoa(registro);
    pessoas[chave] = true;
    atendimentosPorPessoa[chave] = (atendimentosPorPessoa[chave] || 0) + 1;

    totalVinculos += registro.quantidadeVinculos;
    if (registro.quantidadeVinculos > 0) atendimentosComVinculos++;

    const tipo = normalizar(registro.tipoAtendimento);
    if (tipo.includes("familiar") || tipo.includes("grupo")) {
      atendimentosFamiliaresOuGrupo++;
    }
  });

  let retornos = 0;
  let pessoasComRetorno = 0;

  Object.keys(atendimentosPorPessoa).forEach(function(chave) {
    if (atendimentosPorPessoa[chave] > 1) {
      retornos += atendimentosPorPessoa[chave] - 1;
      pessoasComRetorno++;
    }
  });

  return {
    totalAtendimentos: registros.length,
    pessoasDistintas: Object.keys(pessoas).length,
    retornos: retornos,
    pessoasComRetorno: pessoasComRetorno,
    atendimentosComVinculos: atendimentosComVinculos,
    totalVinculos: totalVinculos,
    atendimentosFamiliaresOuGrupo: atendimentosFamiliaresOuGrupo
  };
}

function montarDistribuicoesRelatorio(registros) {
  return {
    porMes: contarPorCampo(registros, "mesCadastro"),
    porNAPS: contarPorCampo(registros, "naps"),
    porTipo: contarPorCampo(registros, "tipoAtendimento"),
    porPostoGraduacao: contarPorCampo(registros, "postoGraduacao"),
    porMotivo: contarPorCampo(registros, "motivo"),
    porOPM: contarPorCampo(registros, "opmAtual"),
    porCidade: contarPorCampo(registros, "cidade"),
    porBairro: contarPorCampo(registros, "bairro"),
    porUF: contarPorCampo(registros, "estado"),
    porResponsavel: contarPorCampo(registros, "responsavel"),
    porSituacao: contarPorCampo(registros, "situacaoStatus"),
    porSexo: contarPorCampo(registros, "sexo"),
    porEstadoCivil: contarPorCampo(registros, "estadoCivil"),
    porFaixaEtaria: contarPorCampo(registros, "faixaEtaria"),
    porTempoServico: contarPorCampo(registros, "tempoServico"),
    porParentesco: contarPorVinculo(registros, "parentesco")
  };
}

function montarDetalhesRelatorio(registros, limite) {
  return registros
    .slice()
    .sort(function(a, b) {
      return b.dataCadastroTimestamp - a.dataCadastroTimestamp;
    })
    .slice(0, limite)
    .map(function(registro) {
      return {
        dataCadastro: registro.dataCadastro,
        dataCadastroIso: registro.dataCadastroIso,
        postoGraduacao: registro.postoGraduacao,
        nome: registro.nome,
        re: registro.re,
        cpf: registro.cpf,
        telefone: registro.telefone,
        tipoAtendimento: registro.tipoAtendimento,
        motivo: registro.motivo,
        naps: registro.naps,
        emailCadastro: registro.emailCadastro,
        opmAtual: registro.opmAtual,
        situacaoStatus: registro.situacaoStatus,
        sexo: registro.sexo,
        idade: registro.idade,
        faixaEtaria: registro.faixaEtaria,
        estadoCivil: registro.estadoCivil,
        responsavel: registro.responsavel,
        cep: registro.cep,
        endereco: registro.endereco,
        bairro: registro.bairro,
        cidade: registro.cidade,
        estado: registro.estado,
        quantidadeVinculos: registro.quantidadeVinculos,
        vinculos: registro.vinculos
      };
    });
}

function montarOpcoesRelatorio(registros, vinculosPorAtendimento) {
  const todosVinculos = [];

  Object.keys(vinculosPorAtendimento).forEach(function(idAtendimento) {
    vinculosPorAtendimento[idAtendimento].forEach(function(vinculo) {
      todosVinculos.push(vinculo);
    });
  });

  return {
    tipoAtendimento: listarUnicosRelatorio(registros, "tipoAtendimento"),
    naps: listarUnicosRelatorio(registros, "naps"),
    opmAtual: listarUnicosRelatorio(registros, "opmAtual"),
    cidade: listarUnicosRelatorio(registros, "cidade"),
    bairro: listarUnicosRelatorio(registros, "bairro"),
    estado: listarUnicosRelatorio(registros, "estado"),
    responsavel: listarUnicosRelatorio(registros, "responsavel"),
    situacaoStatus: listarUnicosRelatorio(registros, "situacaoStatus"),
    sexo: listarUnicosRelatorio(registros, "sexo"),
    estadoCivil: listarUnicosRelatorio(registros, "estadoCivil"),
    parentesco: listarUnicosRelatorio(todosVinculos, "parentesco")
  };
}

function contarPorCampo(registros, campo) {
  const contagem = {};

  registros.forEach(function(registro) {
    adicionarContagemRelatorio(contagem, registro[campo]);
  });

  return contagem;
}

function contarPorVinculo(registros, campo) {
  const contagem = {};

  registros.forEach(function(registro) {
    registro.vinculos.forEach(function(vinculo) {
      adicionarContagemRelatorio(contagem, vinculo[campo]);
    });
  });

  return contagem;
}

function adicionarContagemRelatorio(contagem, valor) {
  const chave = valor || "nao informado";
  contagem[chave] = (contagem[chave] || 0) + 1;
}

function listarUnicosRelatorio(lista, campo) {
  const mapa = {};

  lista.forEach(function(item) {
    const valor = item[campo];
    if (valor !== "" && valor !== null && valor !== undefined) {
      mapa[String(valor)] = true;
    }
  });

  return Object.keys(mapa).sort(function(a, b) {
    return normalizar(a).localeCompare(normalizar(b));
  });
}

function obterChavePessoa(registro) {
  return somenteNumeros(registro.cpf) ||
    normalizar(registro.re) ||
    normalizar(registro.nome) ||
    "id-" + registro.id;
}

function montarEnderecoRelatorio(linha) {
  const partes = [
    linha[18],
    linha[22],
    linha[19],
    linha[20],
    linha[21],
    linha[17]
  ].filter(function(parte) {
    return parte !== "" && parte !== null && parte !== undefined;
  });

  return partes.join(", ");
}

function obterFaixaEtaria(idade) {
  if (idade === null || idade === undefined || idade < 0) return "nao informado";
  if (idade <= 17) return "ate 17 anos";
  if (idade <= 29) return "18 a 29 anos";
  if (idade <= 39) return "30 a 39 anos";
  if (idade <= 49) return "40 a 49 anos";
  if (idade <= 59) return "50 a 59 anos";
  return "60 anos ou mais";
}

function obterFaixaTempoServico(anos) {
  if (anos === null || anos === undefined || anos < 0) return "nao informado";
  if (anos <= 5) return "ate 5 anos";
  if (anos <= 10) return "6 a 10 anos";
  if (anos <= 20) return "11 a 20 anos";
  if (anos <= 30) return "21 a 30 anos";
  return "31 anos ou mais";
}

function calcularAnosAteHoje(data) {
  if (!data) return null;

  const hoje = new Date();
  let anos = hoje.getFullYear() - data.getFullYear();
  const mes = hoje.getMonth() - data.getMonth();

  if (mes < 0 || (mes === 0 && hoje.getDate() < data.getDate())) {
    anos--;
  }

  return anos;
}

function obterMesAno(data) {
  if (!data) return "sem data";

  const ano = data.getFullYear();
  const mes = String(data.getMonth() + 1).padStart(2, "0");

  return ano + "-" + mes;
}

function obterDataInicio(valor) {
  const data = obterData(valor);

  if (!data) return null;

  data.setHours(0, 0, 0, 0);
  return data;
}

function obterDataFim(valor) {
  const data = obterData(valor);

  if (!data) return null;

  data.setHours(23, 59, 59, 999);
  return data;
}

function obterData(valor) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate(), valor.getHours(), valor.getMinutes(), valor.getSeconds());
  }

  const texto = String(valor).trim();
  let partes;

  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) {
    partes = texto.substring(0, 10).split("-");
    return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
  }

  if (/^\d{2}\/\d{2}\/\d{4}/.test(texto)) {
    partes = texto.substring(0, 10).split("/");
    return new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  }

  return null;
}

function obterUrlApp() {
  return ScriptApp.getService().getUrl();
}

function formatarDataBrasil(valor) {
  const data = obterData(valor);

  if (!data) return valor ? String(valor) : "";

  const dia = String(data.getDate()).padStart(2, "0");
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  const ano = data.getFullYear();

  return dia + "/" + mes + "/" + ano;
}

function autorizarServicosSAIC() {
  UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo?id_token=teste", {
    muteHttpExceptions: true
  });

  SpreadsheetApp.getActiveSpreadsheet().getName();
}
