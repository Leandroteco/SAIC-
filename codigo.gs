const ABA_DADOS = "dados_cadastro";
const ABA_VINCULOS = "pessoas_vinculadas";
const ABA_INDICE = "cadastro_indice";
const ABA_USUARIOS = "usuarios_sistema";
const ABA_CEPS_CACHE = "ceps_cache";
const ABA_RECADOS = "recados_sistema";
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

const CABECALHOS_INDICE = [
  "chave",
  "tipo_chave",
  "id_atendimento",
  "cpf",
  "re",
  "nome",
  "dataCadastro",
  "linha_dados"
];

const CABECALHOS_USUARIOS = [
  "email",
  "perfil",
  "ativo",
  "nome",
  "naps",
  "cep_naps",
  "vencimento"
];

const CABECALHOS_CEPS_CACHE = [
  "cep",
  "latitude",
  "longitude",
  "endereco",
  "data_atualizacao"
];

const CABECALHOS_RECADOS = [
  "id_recado",
  "titulo",
  "mensagem",
  "cor",
  "ativo",
  "data_inicio",
  "data_fim",
  "naps_destino",
  "data_criacao"
];

function configurarEstruturaPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDados = obterOuCriarAba(ss, ABA_DADOS, CABECALHOS_DADOS);
  const sheetVinculos = obterOuCriarAba(ss, ABA_VINCULOS, CABECALHOS_VINCULOS);
  const sheetIndice = obterOuCriarAba(ss, ABA_INDICE, CABECALHOS_INDICE);
  const sheetUsuarios = obterOuCriarAba(ss, ABA_USUARIOS, CABECALHOS_USUARIOS);
  const sheetCepsCache = obterOuCriarAba(ss, ABA_CEPS_CACHE, CABECALHOS_CEPS_CACHE);
  const sheetRecados = obterOuCriarAba(ss, ABA_RECADOS, CABECALHOS_RECADOS);

  return {
    sheetDados: sheetDados,
    sheetVinculos: sheetVinculos,
    sheetIndice: sheetIndice,
    sheetUsuarios: sheetUsuarios,
    sheetCepsCache: sheetCepsCache,
    sheetRecados: sheetRecados
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

function obterIndicesUsuariosSistema(sheet) {
  const mapa = obterMapaCabecalhos(sheet);

  return {
    email: obterIndiceCabecalho(mapa, ["email", "e-mail"], 0),
    perfil: obterIndiceCabecalho(mapa, ["perfil"], 1),
    ativo: obterIndiceCabecalho(mapa, ["ativo", "status"], 2),
    nome: obterIndiceCabecalho(mapa, ["nome", "nomeusuario", "nome_usuario"], 3),
    naps: obterIndiceCabecalho(mapa, ["naps", "nap", "unidade", "unidadenaps", "nome_naps"], 4),
    cepNaps: obterIndiceCabecalho(mapa, ["cep_naps", "cepnaps", "cepdonaps", "cep_do_naps", "cep"], 5),
    vencimento: obterIndiceCabecalho(mapa, ["vencimento", "validade", "data_vencimento", "datavencimento"], 6)
  };
}

function obterMapaCabecalhos(sheet) {
  const mapa = {};
  const ultimaColuna = Math.max(sheet.getLastColumn(), CABECALHOS_USUARIOS.length);
  const cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];

  cabecalhos.forEach(function(cabecalho, indice) {
    const chave = normalizarCabecalho(cabecalho);
    if (chave && mapa[chave] === undefined) {
      mapa[chave] = indice;
    }
  });

  return mapa;
}

function normalizarCabecalho(valor) {
  return normalizar(valor).replace(/[^a-z0-9]/g, "");
}

function obterIndiceCabecalho(mapa, aliases, fallback) {
  for (let i = 0; i < aliases.length; i++) {
    const chave = normalizarCabecalho(aliases[i]);
    if (mapa[chave] !== undefined) return mapa[chave];
  }

  return fallback;
}

function lerDadosUsuarioSistema(linha, indices) {
  return {
    email: String(linha[indices.email] || "").toLowerCase().trim(),
    perfil: String(linha[indices.perfil] || "").toLowerCase().trim(),
    ativo: String(linha[indices.ativo] || "").toLowerCase().trim(),
    nome: String(linha[indices.nome] || "").trim(),
    naps: String(linha[indices.naps] || "").trim(),
    cepNaps: formatarCEP(linha[indices.cepNaps] || ""),
    vencimento: linha[indices.vencimento] || ""
  };
}

function converterDataVencimentoUsuario(valor) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }

  const texto = String(valor || "").trim();
  if (!texto) return null;

  let partes = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (partes) {
    return new Date(Number(partes[3]), Number(partes[2]) - 1, Number(partes[1]));
  }

  partes = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (partes) {
    return new Date(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3]));
  }

  return null;
}

function obterInicioHoje() {
  const hoje = new Date();
  return new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
}

function usuarioSistemaVencido(usuario) {
  const vencimento = converterDataVencimentoUsuario(usuario.vencimento);
  if (!vencimento) return false;

  return vencimento < obterInicioHoje();
}

function formatarDataVencimentoUsuario(valor) {
  const vencimento = converterDataVencimentoUsuario(valor);
  if (!vencimento) return "";

  return Utilities.formatDate(vencimento, Session.getScriptTimeZone(), "dd/MM/yyyy");
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
      naps: "",
      cepNaps: "",
      mensagem: "Nao foi possivel identificar o email Google do usuario. Verifique a implantacao do Web App."
    };
  }

  const dados = sheet.getDataRange().getValues();
  const indices = obterIndicesUsuariosSistema(sheet);

  for (let i = 1; i < dados.length; i++) {
    const usuarioLinha = lerDadosUsuarioSistema(dados[i], indices);

    if (usuarioLinha.email === email) {
      if (usuarioLinha.ativo !== "sim") {
        return {
          autorizado: false,
          email: email,
          perfil: usuarioLinha.perfil,
          nome: usuarioLinha.nome,
          naps: usuarioLinha.naps,
          cepNaps: usuarioLinha.cepNaps,
          vencimento: usuarioLinha.vencimento,
          mensagem: "Usuario cadastrado, mas inativo."
        };
      }

      if (usuarioSistemaVencido(usuarioLinha)) {
        return {
          autorizado: false,
          email: email,
          perfil: usuarioLinha.perfil,
          nome: usuarioLinha.nome,
          naps: usuarioLinha.naps,
          cepNaps: usuarioLinha.cepNaps,
          vencimento: usuarioLinha.vencimento,
          mensagem: "Usuario com acesso vencido em " + formatarDataVencimentoUsuario(usuarioLinha.vencimento) + "."
        };
      }

      return {
        autorizado: true,
        email: email,
        perfil: usuarioLinha.perfil,
        nome: usuarioLinha.nome,
        naps: usuarioLinha.naps,
        cepNaps: usuarioLinha.cepNaps,
        vencimento: usuarioLinha.vencimento,
        mensagem: ""
      };
    }
  }

  return {
    autorizado: false,
    email: email,
    perfil: "",
    nome: "",
    naps: "",
    cepNaps: "",
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
  const pagina = obterPaginaSolicitada(e && e.parameter ? e.parameter.pagina : "");

  return renderizarPagina(pagina, "", null, "");
}

function obterPaginaSolicitada(valor) {
  const pagina = String(valor || "").toLowerCase().trim();

  if (pagina === "relatorios") return "relatorios";
  if (pagina === "dashboard") return "dashboard";
  if (pagina === "mapa_de_calor") return "mapa_de_calor";

  return "Index";
}

function paginaExigeAdministrador(pagina) {
  return pagina === "relatorios" || pagina === "dashboard" || pagina === "mapa_de_calor";
}

function obterConfigLoginGoogle() {
  return {
    clientId: GOOGLE_CLIENT_ID
  };
}

function doPost(e) {
  const idToken = e && e.parameter ? String(e.parameter.credential || "") : "";
  const destino = e && e.parameter ? String(e.parameter.state || "index") : "index";
  const pagina = obterPaginaSolicitada(destino);

  try {
    const usuario = validarLoginGoogle(idToken);

    if (paginaExigeAdministrador(pagina) && usuario.perfil !== "administrador") {
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
    naps: usuario.naps || "",
    vencimento: usuario.vencimento || "",
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
      naps: "",
      cepNaps: "",
      mensagem: "Email não identificado."
    };
  }

  const dados = sheet.getDataRange().getValues();
  const indices = obterIndicesUsuariosSistema(sheet);

  for (let i = 1; i < dados.length; i++) {
    const usuarioLinha = lerDadosUsuarioSistema(dados[i], indices);

    if (usuarioLinha.email === emailNormalizado) {
      if (usuarioLinha.ativo !== "sim") {
        return {
          autorizado: false,
          email: emailNormalizado,
          perfil: usuarioLinha.perfil,
          nome: usuarioLinha.nome,
          naps: usuarioLinha.naps,
          cepNaps: usuarioLinha.cepNaps,
          vencimento: usuarioLinha.vencimento,
          mensagem: "Usuário cadastrado, mas inativo."
        };
      }

      if (usuarioSistemaVencido(usuarioLinha)) {
        return {
          autorizado: false,
          email: emailNormalizado,
          perfil: usuarioLinha.perfil,
          nome: usuarioLinha.nome,
          naps: usuarioLinha.naps,
          cepNaps: usuarioLinha.cepNaps,
          vencimento: usuarioLinha.vencimento,
          mensagem: "Usuário com acesso vencido em " + formatarDataVencimentoUsuario(usuarioLinha.vencimento) + "."
        };
      }

      return {
        autorizado: true,
        email: emailNormalizado,
        perfil: usuarioLinha.perfil,
        nome: usuarioLinha.nome,
        naps: usuarioLinha.naps,
        cepNaps: usuarioLinha.cepNaps,
        vencimento: usuarioLinha.vencimento,
        mensagem: ""
      };
    }
  }

  return {
    autorizado: false,
    email: emailNormalizado,
    perfil: "",
    nome: "",
    naps: "",
    cepNaps: "",
    mensagem: "Email não autorizado: " + emailNormalizado
  };
}

function obterRecadoAtivo(idToken) {
  const usuario = validarUsuarioPorToken(idToken);
  const estrutura = configurarEstruturaPlanilha();
  const sheet = estrutura.sheetRecados;

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      encontrado: false
    };
  }

  const dados = sheet.getDataRange().getValues();
  const indices = obterIndicesRecadosSistema(sheet);
  const hoje = new Date();
  const recadosValidos = [];

  for (let i = 1; i < dados.length; i++) {
    const recado = lerRecadoSistema(dados[i], indices);

    if (!recado.id && !recado.titulo && !recado.mensagem) continue;
    if (normalizar(recado.ativo) !== "sim") continue;
    if (!recado.mensagem) continue;
    if (!recadoDentroDoPeriodo(recado, hoje)) continue;
    if (!recadoDestinadoAoNaps(recado, usuario)) continue;

    recadosValidos.push(recado);
  }

  if (recadosValidos.length === 0) {
    return {
      encontrado: false
    };
  }

  recadosValidos.sort(function(a, b) {
    return b.dataCriacaoTimestamp - a.dataCriacaoTimestamp;
  });

  const recadoSelecionado = recadosValidos[0];

  return {
    encontrado: true,
    id: recadoSelecionado.id,
    titulo: recadoSelecionado.titulo || "Recado",
    mensagem: recadoSelecionado.mensagem,
    cor: normalizarCorRecado(recadoSelecionado.cor),
    dataInicio: formatarDataBrasil(recadoSelecionado.dataInicio),
    dataFim: formatarDataBrasil(recadoSelecionado.dataFim)
  };
}

function obterIndicesRecadosSistema(sheet) {
  const mapa = obterMapaCabecalhos(sheet);

  return {
    id: obterIndiceCabecalho(mapa, ["id_recado", "idrecado", "id"], 0),
    titulo: obterIndiceCabecalho(mapa, ["titulo", "title"], 1),
    mensagem: obterIndiceCabecalho(mapa, ["mensagem", "recado", "texto"], 2),
    cor: obterIndiceCabecalho(mapa, ["cor", "color"], 3),
    ativo: obterIndiceCabecalho(mapa, ["ativo", "status"], 4),
    dataInicio: obterIndiceCabecalho(mapa, ["data_inicio", "datainicio", "inicio"], 5),
    dataFim: obterIndiceCabecalho(mapa, ["data_fim", "datafim", "fim"], 6),
    napsDestino: obterIndiceCabecalho(mapa, ["naps_destino", "napsdestino", "destino", "naps"], 7),
    dataCriacao: obterIndiceCabecalho(mapa, ["data_criacao", "datacriacao", "criacao"], 8)
  };
}

function lerRecadoSistema(linha, indices) {
  const dataInicio = obterData(linha[indices.dataInicio]);
  const dataFim = obterData(linha[indices.dataFim]);
  const dataCriacao = obterData(linha[indices.dataCriacao]);

  return {
    id: String(linha[indices.id] || "").trim(),
    titulo: String(linha[indices.titulo] || "").trim(),
    mensagem: String(linha[indices.mensagem] || "").trim(),
    cor: String(linha[indices.cor] || "").trim(),
    ativo: String(linha[indices.ativo] || "").trim(),
    dataInicio: dataInicio,
    dataFim: dataFim,
    napsDestino: String(linha[indices.napsDestino] || "").trim(),
    dataCriacao: dataCriacao,
    dataCriacaoTimestamp: dataCriacao ? dataCriacao.getTime() : 0
  };
}

function recadoDentroDoPeriodo(recado, hoje) {
  const agora = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const dataInicio = recado.dataInicio
    ? new Date(recado.dataInicio.getFullYear(), recado.dataInicio.getMonth(), recado.dataInicio.getDate())
    : null;
  const dataFim = recado.dataFim
    ? new Date(recado.dataFim.getFullYear(), recado.dataFim.getMonth(), recado.dataFim.getDate())
    : null;

  if (dataInicio && agora < dataInicio) return false;
  if (dataFim && agora > dataFim) return false;

  return true;
}

function recadoDestinadoAoNaps(recado, usuario) {
  const destino = normalizar(recado.napsDestino || "todos");

  if (!destino || destino === "todos") return true;

  return normalizar(usuario.naps) === destino;
}

function normalizarCorRecado(cor) {
  const corNormalizada = normalizar(cor || "amarelo");
  const coresPermitidas = {
    amarelo: true,
    azul: true,
    verde: true,
    rosa: true
  };

  return coresPermitidas[corNormalizada] ? corNormalizada : "amarelo";
}


function salvarAtendimento(dados, idToken) {
  const usuario = validarUsuarioPorToken(idToken);
  const lock = LockService.getScriptLock();
  let lockObtido = false;

  try {
    lock.waitLock(30000);
    lockObtido = true;

    const estrutura = configurarEstruturaPlanilha();
    const sheet = estrutura.sheetDados;
    const sheetVinculos = estrutura.sheetVinculos;
    const sheetIndice = estrutura.sheetIndice;

    if (!sheet) throw new Error("A aba dados_cadastro nao foi encontrada.");
    if (!sheetVinculos) throw new Error("A aba pessoas_vinculadas nao foi encontrada.");
    if (!sheetIndice) throw new Error("A aba cadastro_indice nao foi encontrada.");

    const idAtendimento = gerarIdSeguro("ATD");
    const dataCadastro = new Date();
    const dataIngresso = validarDataFormulario(dados.dataIngresso, "Data de Ingresso");
    const dataNascimento = validarDataFormulario(dados.dataNascimento, "Data de Nascimento");
    const dataInatividade = validarDataFormulario(dados.dataInatividade, "Data de Inatividade");
    const tipoAtendimento = normalizarTipoAtendimento(dados.tipoAtendimento);
    const responsavelAtendimento = normalizar(usuario.nome || dados.responsavel);

    if (!tipoAtendimento) {
      throw new Error("Tipo de atendimento e obrigatorio.");
    }

    if (!responsavelAtendimento) {
      throw new Error("Nome do responsavel nao encontrado na aba usuarios_sistema.");
    }

    sheet.appendRow([
      idAtendimento,
      usuario.email,
      tipoAtendimento,
      normalizar(dados.motivo),
      normalizar(dados.re),
      normalizar(dados.nome),
      formatarCPF(dados.cpf),
      normalizarTelefone(dados.telefone),
      normalizar(dados.email),
      dataIngresso,
      dataNascimento,
      normalizar(dados.sexo),
      normalizar(dados.opmAtual),
      normalizar(dados.situacaoStatus),
      dataInatividade,
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
      responsavelAtendimento,
      dataCadastro,
      dados.postoGraduacao || ""
    ]);

    const linhaDados = sheet.getLastRow();
    const linhasIndice = montarLinhasIndiceCadastro(idAtendimento, dados, dataCadastro, linhaDados);

    if (linhasIndice.length > 0) {
      gravarLinhasAbaixo(sheetIndice, linhasIndice, CABECALHOS_INDICE.length);
    }

    const linhasVinculos = [];

    if (dados.pessoasVinculadas && dados.pessoasVinculadas.length > 0) {
      dados.pessoasVinculadas.forEach(function(pessoa) {
        if (pessoa.nome || pessoa.cpf) {
          linhasVinculos.push([
            gerarIdSeguro("VIN"),
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

    if (linhasVinculos.length > 0) {
      gravarLinhasAbaixo(sheetVinculos, linhasVinculos, CABECALHOS_VINCULOS.length);
    }

    return tipoAtendimento === "falta"
      ? "Falta registrada com sucesso!"
      : "Atendimento salvo com sucesso!";
  } finally {
    if (lockObtido) {
      lock.releaseLock();
    }
  }
}

function gerarIdSeguro(prefixo) {
  const id = Utilities.getUuid();

  return prefixo ? prefixo + "-" + id : id;
}

function validarDataFormulario(valor, nomeCampo) {
  const data = String(valor || "").trim();

  if (!data) return "";

  if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) {
    throw new Error(nomeCampo + " deve ter ano com 4 digitos.");
  }

  const ano = Number(data.substring(0, 4));
  const mes = Number(data.substring(5, 7));
  const dia = Number(data.substring(8, 10));

  if (ano < 1 || ano > 9999 || mes < 1 || mes > 12) {
    throw new Error(nomeCampo + " invalida.");
  }

  const anoBissexto = (ano % 4 === 0 && ano % 100 !== 0) || ano % 400 === 0;
  const diasPorMes = [31, anoBissexto ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  if (dia < 1 || dia > diasPorMes[mes - 1]) {
    throw new Error(nomeCampo + " invalida.");
  }

  return data;
}

function montarLinhasIndiceCadastro(idAtendimento, dados, dataCadastro, linhaDados) {
  const cpf = formatarCPF(dados.cpf);
  const cpfNumeros = somenteNumeros(cpf);
  const re = normalizar(dados.re);
  const reNumeros = somenteNumeros(re);
  const nome = normalizar(dados.nome);
  const linhas = [];

  if (cpfNumeros) {
    linhas.push([cpfNumeros, "cpf", idAtendimento, cpf, re, nome, dataCadastro, linhaDados]);
  }

  if (reNumeros) {
    linhas.push([reNumeros, "re", idAtendimento, cpf, re, nome, dataCadastro, linhaDados]);
  }

  if (nome) {
    linhas.push([nome, "nome", idAtendimento, cpf, re, nome, dataCadastro, linhaDados]);
  }

  return linhas;
}

function gravarLinhasAbaixo(sheet, linhas, quantidadeColunas) {
  const tamanhoLote = 5000;

  for (let i = 0; i < linhas.length; i += tamanhoLote) {
    const lote = linhas.slice(i, i + tamanhoLote);

    sheet
      .getRange(sheet.getLastRow() + 1, 1, lote.length, quantidadeColunas)
      .setValues(lote);
  }
}

function reconstruirIndiceCadastros() {
  const estrutura = configurarEstruturaPlanilha();
  const sheetDados = estrutura.sheetDados;
  const sheetIndice = estrutura.sheetIndice;

  if (!sheetDados) throw new Error("A aba dados_cadastro nao foi encontrada.");
  if (!sheetIndice) throw new Error("A aba cadastro_indice nao foi encontrada.");

  if (sheetIndice.getLastRow() > 1) {
    sheetIndice
      .getRange(2, 1, sheetIndice.getLastRow() - 1, CABECALHOS_INDICE.length)
      .clearContent();
  }

  if (sheetDados.getLastRow() < 2) {
    return "Indice reconstruido. Nenhum cadastro encontrado.";
  }

  const dados = sheetDados
    .getRange(2, 1, sheetDados.getLastRow() - 1, CABECALHOS_DADOS.length)
    .getValues();
  const linhasIndice = [];

  dados.forEach(function(linha, indice) {
    const dadosCadastro = {
      cpf: linha[6],
      re: linha[4],
      nome: linha[5]
    };

    Array.prototype.push.apply(
      linhasIndice,
      montarLinhasIndiceCadastro(linha[0], dadosCadastro, linha[26], indice + 2)
    );
  });

  if (linhasIndice.length > 0) {
    gravarLinhasAbaixo(sheetIndice, linhasIndice, CABECALHOS_INDICE.length);
  }

  return "Indice reconstruido com sucesso. Chaves criadas: " + linhasIndice.length + ".";
}

function normalizar(texto) {
  if (!texto) return "";

  return String(texto)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function normalizarTipoAtendimento(valor) {
  const tipo = normalizar(valor);

  if (tipo === "alta" || tipo === "atendimento de alta") return "alta";
  if (tipo === "falta" || tipo === "registrar falta" || tipo.includes("no show")) return "falta";

  return tipo;
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

function formatarCEP(cep) {
  cep = somenteNumeros(cep);

  if (cep.length !== 8) return cep;

  return cep.replace(/(\d{5})(\d{3})/, "$1-$2");
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_DADOS);
  const sheetIndice = obterOuCriarAba(ss, ABA_INDICE, CABECALHOS_INDICE);

  if (!sheet) {
    return {
      encontrado: false,
      mensagem: "A aba dados_cadastro nao foi encontrada."
    };
  }

  const buscaTexto = normalizar(termo);
  const buscaNumeros = somenteNumeros(termo);
  const termoOriginal = String(termo || "").trim();

  if (buscaTexto.length < 5 && buscaNumeros.length < 5) {
    return {
      encontrado: false,
      mensagem: "Digite pelo menos 5 caracteres para R.E./Nome ou informe o CPF completo."
    };
  }

  if (buscaNumeros.length >= 8 && buscaNumeros.length < 11 && !/[A-Z]/i.test(termoOriginal)) {
    return {
      encontrado: false,
      mensagem: "Para pesquisar por CPF, informe o CPF completo."
    };
  }

  const pesquisouCPFCompleto = buscaNumeros.length === 11;
  const pesquisouRECompleto = /^[0-9]{6}-[0-9A]$/i.test(termoOriginal);

  const candidatos = localizarCadastrosNoIndice(
    sheetIndice,
    buscaTexto,
    buscaNumeros,
    pesquisouCPFCompleto
  );
  const resultados = montarResultadosBuscaPorIndice(sheet, candidatos);

  if (resultados.length === 0) {
    return {
      encontrado: false,
      mensagem: "Nenhum cadastro anterior localizado. Preencha novo cadastro."
    };
  }

  if (pesquisouCPFCompleto || pesquisouRECompleto) {
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

function localizarCadastrosNoIndice(sheetIndice, buscaTexto, buscaNumeros, pesquisouCPFCompleto) {
  const lastRow = sheetIndice.getLastRow();

  if (lastRow < 2) return [];

  const linhas = sheetIndice.getRange(2, 1, lastRow - 1, CABECALHOS_INDICE.length).getValues();
  const encontrados = [];

  for (let i = 0; i < linhas.length; i++) {
    const linha = linhas[i];
    const chave = String(linha[0] || "");
    const tipoChave = String(linha[1] || "");
    const linhaDados = Number(linha[7] || 0);

    if (!linhaDados) continue;

    const achouCPF = tipoChave === "cpf" && pesquisouCPFCompleto && chave === buscaNumeros;
    const achouRE = tipoChave === "re" && buscaNumeros.length >= 5 && chave.includes(buscaNumeros);
    const achouNome = tipoChave === "nome" && buscaTexto.length >= 5 && chave.includes(buscaTexto);

    if (achouCPF || achouRE || achouNome) {
      encontrados.push({
        idAtendimento: linha[2],
        cpf: linha[3],
        re: linha[4],
        nome: linha[5],
        dataCadastro: linha[6],
        linhaDados: linhaDados
      });
    }
  }

  encontrados.sort(function(a, b) {
    return b.linhaDados - a.linhaDados;
  });

  return encontrados;
}

function montarResultadosBuscaPorIndice(sheetDados, candidatos) {
  const resultados = [];
  const chavesEncontradas = {};

  for (let i = 0; i < candidatos.length; i++) {
    const candidato = candidatos[i];
    const chaveUnica = somenteNumeros(candidato.cpf) ||
      normalizar(candidato.re) ||
      normalizar(candidato.nome) ||
      String(candidato.idAtendimento || "");

    if (chavesEncontradas[chaveUnica]) continue;
    chavesEncontradas[chaveUnica] = true;

    const linha = sheetDados
      .getRange(candidato.linhaDados, 1, 1, CABECALHOS_DADOS.length)
      .getValues()[0];

    resultados.push(montarRegistroBusca(linha));

    if (resultados.length >= 10) break;
  }

  return resultados;
}

function montarRegistroBusca(l) {
  return {
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
  const usuariosPorEmail = carregarUsuariosSistemaPorEmail(ss);
  const registros = [];

  for (let i = 1; i < dados.length; i++) {
    registros.push(montarRegistroRelatorio(dados[i], vinculosPorAtendimento, usuariosPorEmail));
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
    dadosIndividuais: montarDadosIndividuaisRelatorio(filtros, filtrados, registros, ss),
    registros: montarDetalhesRelatorio(filtrados, limiteDetalhes),
    totalRegistros: filtrados.length,
    limiteDetalhes: limiteDetalhes,
    opcoes: montarOpcoesRelatorio(registros, vinculosPorAtendimento)
  };
}

function obterDadosDashboard(filtros, idToken) {
  validarAdministradorPorToken(idToken);

  const filtrosDashboard = normalizarFiltrosDashboard(filtros || {});
  const estrutura = configurarEstruturaPlanilha();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = estrutura.sheetDados;
  const dados = sheet.getDataRange().getValues();
  const vinculosPorAtendimento = carregarVinculosPorAtendimento(ss);
  const usuariosPorEmail = carregarUsuariosSistemaPorEmail(ss);
  const registros = [];

  for (let i = 1; i < dados.length; i++) {
    registros.push(montarRegistroRelatorio(dados[i], vinculosPorAtendimento, usuariosPorEmail));
  }

  const filtrosPreparados = prepararFiltrosRelatorio(filtrosDashboard);
  const filtrados = registros.filter(function(registro) {
    return registroPassaFiltrosRelatorio(registro, filtrosPreparados);
  });

  return {
    filtrosAplicados: filtrosDashboard,
    periodo: montarPeriodoDashboard(filtrosDashboard),
    resumo: montarResumoDashboard(filtrados, filtrosDashboard),
    graficos: montarGraficosDashboard(filtrados, filtrosDashboard),
    rankings: montarRankingsDashboard(filtrados),
    opcoes: montarOpcoesRelatorio(registros, vinculosPorAtendimento),
    totalRegistros: filtrados.length
  };
}

function obterDadosMapaCalor(filtros, idToken) {
  validarAdministradorPorToken(idToken);

  const filtrosMapa = normalizarFiltrosMapaCalor(filtros || {});
  const estrutura = configurarEstruturaPlanilha();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = estrutura.sheetDados;
  const dados = sheet.getDataRange().getValues();
  const usuariosPorEmail = carregarUsuariosSistemaPorEmail(ss);
  const registros = [];

  for (let i = 1; i < dados.length; i++) {
    registros.push(montarRegistroRelatorio(dados[i], {}, usuariosPorEmail));
  }

  const filtrosPreparados = prepararFiltrosRelatorio(filtrosMapa);
  const filtrados = registros.filter(function(registro) {
    return registroPassaFiltrosRelatorio(registro, filtrosPreparados) && !ehRegistroFalta(registro);
  });

  const napsReferencia = carregarNapsReferenciaMapaCalor(ss, usuariosPorEmail);
  const contextoCoordenadas = criarContextoCoordenadasMapa(ss);
  const gruposRegiao = {};
  const cargaNaps = {};
  const pessoasDistintas = {};
  const distanciaAlerta = Number(filtrosMapa.distanciaAlertaKm || 20);
  let totalDistanciaAtual = 0;
  let totalComDistancia = 0;
  let cepsSemCadastro = 0;
  let cepsSemCoordenada = 0;
  let registrosSemNaps = 0;

  filtrados.forEach(function(registro) {
    const chavePessoa = obterChavePessoa(registro);
    const cepResidencia = somenteNumeros(registro.cep);
    const nomeNaps = String(registro.naps || "nao informado").trim() || "nao informado";
    const chaveNaps = normalizar(nomeNaps);
    const dadosNaps = napsReferencia[chaveNaps] || null;

    pessoasDistintas[chavePessoa] = true;

    if (cepResidencia.length !== 8) {
      cepsSemCadastro++;
      if (!dadosNaps || !dadosNaps.coordenadas) registrosSemNaps++;
      atualizarNapsCargaMapa(cargaNaps, nomeNaps, dadosNaps, registro, null, distanciaAlerta);
      return;
    }

    const coordenadasResidencia = obterCoordenadasCEPMapa(cepResidencia, contextoCoordenadas);

    if (!coordenadasResidencia) {
      cepsSemCoordenada++;
      if (!dadosNaps || !dadosNaps.coordenadas) registrosSemNaps++;
      atualizarNapsCargaMapa(cargaNaps, nomeNaps, dadosNaps, registro, null, distanciaAlerta);
      return;
    }

    let distanciaAtual = null;

    if (dadosNaps && dadosNaps.coordenadas) {
      distanciaAtual = calcularDistanciaHaversineKm(
        coordenadasResidencia.latitude,
        coordenadasResidencia.longitude,
        dadosNaps.coordenadas.latitude,
        dadosNaps.coordenadas.longitude
      );
      totalDistanciaAtual += distanciaAtual;
      totalComDistancia++;
    } else {
      registrosSemNaps++;
    }

    atualizarNapsCargaMapa(cargaNaps, nomeNaps, dadosNaps, registro, distanciaAtual, distanciaAlerta);

    const cepBase = formatarCepBaseMapa(cepResidencia);
    const chaveRegiao = cepBase || cepResidencia;

    if (!gruposRegiao[chaveRegiao]) {
      gruposRegiao[chaveRegiao] = criarGrupoRegiaoMapa(chaveRegiao, cepBase);
    }

    const grupo = gruposRegiao[chaveRegiao];
    grupo.atendimentos++;
    grupo.pessoas[chavePessoa] = true;
    grupo.somaLat += coordenadasResidencia.latitude;
    grupo.somaLng += coordenadasResidencia.longitude;
    grupo.registros.push({
      latitude: coordenadasResidencia.latitude,
      longitude: coordenadasResidencia.longitude,
      distanciaAtual: distanciaAtual
    });
  });

  salvarNovasCoordenadasMapa(contextoCoordenadas);

  const regioes = montarRegioesMapaCalor(gruposRegiao, filtrosMapa);
  const naps = montarNapsMapaCalor(cargaNaps);
  const resumo = montarResumoMapaCalor({
    totalAtendimentos: filtrados.length,
    pessoasDistintas: Object.keys(pessoasDistintas).length,
    regioesMapeadas: regioes.length,
    distanciaMediaAtualKm: mediaArredondadaMapa(totalDistanciaAtual, totalComDistancia),
    melhorGanhoMedioKm: regioes.length ? regioes[0].ganhoMedioKm : 0,
    cepsSemCadastro: cepsSemCadastro,
    cepsSemCoordenada: cepsSemCoordenada,
    registrosSemNaps: registrosSemNaps,
    cepsGeocodificadosAgora: contextoCoordenadas.geocodificadosAgora,
    limiteGeocodificacaoAtingido: contextoCoordenadas.limiteAtingido
  });

  return {
    filtrosAplicados: filtrosMapa,
    periodo: montarPeriodoDashboard(filtrosMapa),
    resumo: resumo,
    mapa: {
      centro: { latitude: -22.25, longitude: -48.55, zoom: 7 },
      calor: regioes.slice(0, 120),
      naps: naps,
      candidatos: regioes.slice(0, 8)
    },
    rankings: {
      napsSobrecarga: naps.slice(0, 15),
      regioesCandidatas: regioes.slice(0, 15)
    },
    avisos: montarAvisosMapaCalor(resumo)
  };
}

function normalizarFiltrosDashboard(filtros) {
  const hoje = new Date();
  const primeiroDiaMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const ultimoDiaMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);

  return {
    dataInicial: filtros.dataInicial || formatarDataParaInput(primeiroDiaMes),
    dataFinal: filtros.dataFinal || formatarDataParaInput(ultimoDiaMes),
    tipoAtendimento: filtros.tipoAtendimento || "",
    motivo: filtros.motivo || "",
    opmAtual: filtros.opmAtual || "",
    situacaoStatus: filtros.situacaoStatus || "",
    sexo: filtros.sexo || "",
    responsavel: filtros.responsavel || "",
    faixaEtaria: filtros.faixaEtaria || ""
  };
}

function normalizarFiltrosMapaCalor(filtros) {
  const base = normalizarFiltrosDashboard(filtros || {});
  const minimo = Number(filtros.minAtendimentosRegiao || filtros.minAtendimentos || 3);
  const distanciaAlertaKm = Number(filtros.distanciaAlertaKm || 20);

  base.minAtendimentosRegiao = minimo > 0 ? minimo : 3;
  base.distanciaAlertaKm = distanciaAlertaKm > 0 ? distanciaAlertaKm : 20;

  return base;
}

function carregarNapsReferenciaMapaCalor(ss, usuariosPorEmail) {
  const mapa = {};

  Object.keys(usuariosPorEmail || {}).forEach(function(email) {
    const usuario = usuariosPorEmail[email];
    adicionarNapsReferenciaMapa(mapa, usuario.naps, usuario.cepNaps, null);
  });

  const sheetNaps = ss.getSheetByName("naps");

  if (sheetNaps && sheetNaps.getLastRow() > 1) {
    const dados = sheetNaps.getDataRange().getValues();
    const mapaCabecalhos = obterMapaCabecalhos(sheetNaps);
    const indiceNaps = obterIndiceCabecalho(mapaCabecalhos, ["naps", "nap", "nome_naps", "unidade"], 0);
    const indiceCep = obterIndiceCabecalho(mapaCabecalhos, ["cep_naps", "cepnaps", "cep", "cep_do_naps"], 1);

    for (let i = 1; i < dados.length; i++) {
      adicionarNapsReferenciaMapa(mapa, dados[i][indiceNaps], dados[i][indiceCep], null);
    }
  }

  const contexto = criarContextoCoordenadasMapa(ss);

  Object.keys(mapa).forEach(function(chave) {
    const item = mapa[chave];
    item.coordenadas = obterCoordenadasCEPMapa(item.cepNaps, contexto);

    if (item.coordenadas) {
      item.latitude = item.coordenadas.latitude;
      item.longitude = item.coordenadas.longitude;
    }
  });

  salvarNovasCoordenadasMapa(contexto);

  return mapa;
}

function adicionarNapsReferenciaMapa(mapa, nomeNaps, cepNaps, coordenadas) {
  const nome = String(nomeNaps || "").trim();
  const cep = somenteNumeros(cepNaps);

  if (!nome || nome === "nao informado" || cep.length !== 8) return;

  const chave = normalizar(nome);

  if (!mapa[chave]) {
    mapa[chave] = {
      naps: nome,
      cepNaps: cep,
      coordenadas: coordenadas || null,
      latitude: null,
      longitude: null
    };
  }
}

function criarContextoCoordenadasMapa(ss) {
  const sheet = obterOuCriarAba(ss || SpreadsheetApp.getActiveSpreadsheet(), ABA_CEPS_CACHE, CABECALHOS_CEPS_CACHE);
  const dados = sheet.getDataRange().getValues();
  const mapa = {};

  for (let i = 1; i < dados.length; i++) {
    const cep = somenteNumeros(dados[i][0]);
    const latitude = Number(String(dados[i][1] || "").replace(",", "."));
    const longitude = Number(String(dados[i][2] || "").replace(",", "."));

    if (cep.length === 8 && !isNaN(latitude) && !isNaN(longitude)) {
      mapa[cep] = {
        latitude: latitude,
        longitude: longitude
      };
    }
  }

  return {
    sheet: sheet,
    mapa: mapa,
    novasLinhas: [],
    geocodificadosAgora: 0,
    limiteGeocodificacao: 120,
    limiteAtingido: false
  };
}

function obterCoordenadasCEPMapa(cep, contexto) {
  const cepNormalizado = somenteNumeros(cep);

  if (cepNormalizado.length !== 8) return null;
  if (contexto.mapa[cepNormalizado]) return contexto.mapa[cepNormalizado];

  if (contexto.geocodificadosAgora >= contexto.limiteGeocodificacao) {
    contexto.limiteAtingido = true;
    return null;
  }

  const coordenadas = geocodificarCEPMapa(cepNormalizado);

  if (!coordenadas) return null;

  contexto.mapa[cepNormalizado] = coordenadas;
  contexto.geocodificadosAgora++;
  contexto.novasLinhas.push([
    formatarCEP(cepNormalizado),
    coordenadas.latitude,
    coordenadas.longitude,
    coordenadas.endereco || formatarCEP(cepNormalizado) + ", Brasil",
    new Date()
  ]);

  return coordenadas;
}

function geocodificarCEPMapa(cepNormalizado) {
  const endereco = formatarCEP(cepNormalizado) + ", Brasil";
  const resposta = Maps.newGeocoder()
    .setRegion("br")
    .geocode(endereco);

  if (!resposta || resposta.status !== "OK" || !resposta.results || resposta.results.length === 0) {
    return null;
  }

  const localizacao = resposta.results[0].geometry && resposta.results[0].geometry.location;

  if (!localizacao) return null;

  const latitude = Number(localizacao.lat);
  const longitude = Number(localizacao.lng);

  if (isNaN(latitude) || isNaN(longitude)) return null;

  return {
    latitude: latitude,
    longitude: longitude,
    endereco: resposta.results[0].formatted_address || endereco
  };
}

function salvarNovasCoordenadasMapa(contexto) {
  if (!contexto || !contexto.novasLinhas || contexto.novasLinhas.length === 0) return;

  gravarLinhasAbaixo(contexto.sheet, contexto.novasLinhas, CABECALHOS_CEPS_CACHE.length);
}

function criarGrupoRegiaoMapa(chave, cepBase) {
  return {
    chave: chave,
    cepBase: cepBase || chave,
    atendimentos: 0,
    pessoas: {},
    somaLat: 0,
    somaLng: 0,
    registros: []
  };
}

function formatarCepBaseMapa(cep) {
  const numeros = somenteNumeros(cep);

  if (numeros.length !== 8) return "";

  return numeros.substring(0, 5) + "-000";
}

function atualizarNapsCargaMapa(mapa, nomeNaps, dadosNaps, registro, distanciaAtual, distanciaAlertaKm) {
  const nome = String(nomeNaps || "nao informado").trim() || "nao informado";
  const chave = normalizar(nome);

  if (!mapa[chave]) {
    mapa[chave] = {
      naps: nome,
      cepNaps: dadosNaps ? formatarCEP(dadosNaps.cepNaps) : "",
      latitude: dadosNaps && dadosNaps.coordenadas ? dadosNaps.coordenadas.latitude : null,
      longitude: dadosNaps && dadosNaps.coordenadas ? dadosNaps.coordenadas.longitude : null,
      atendimentos: 0,
      pessoas: {},
      somaDistancia: 0,
      totalComDistancia: 0,
      distantes: 0
    };
  }

  mapa[chave].atendimentos++;
  mapa[chave].pessoas[obterChavePessoa(registro)] = true;

  if (distanciaAtual !== null && distanciaAtual !== undefined && !isNaN(distanciaAtual)) {
    mapa[chave].somaDistancia += distanciaAtual;
    mapa[chave].totalComDistancia++;
    if (distanciaAtual >= Number(distanciaAlertaKm || 20)) mapa[chave].distantes++;
  }
}

function montarRegioesMapaCalor(gruposRegiao, filtros) {
  const regioes = [];
  const minimo = Number(filtros.minAtendimentosRegiao || 3);

  Object.keys(gruposRegiao).forEach(function(chave) {
    const grupo = gruposRegiao[chave];

    if (!grupo.atendimentos || grupo.atendimentos < minimo) return;

    const latitude = grupo.somaLat / grupo.atendimentos;
    const longitude = grupo.somaLng / grupo.atendimentos;
    let somaDistanciaAtual = 0;
    let somaDistanciaNova = 0;
    let somaGanho = 0;
    let totalComparavel = 0;

    grupo.registros.forEach(function(registro) {
      if (registro.distanciaAtual === null || registro.distanciaAtual === undefined || isNaN(registro.distanciaAtual)) {
        return;
      }

      const distanciaNova = calcularDistanciaHaversineKm(
        registro.latitude,
        registro.longitude,
        latitude,
        longitude
      );

      somaDistanciaAtual += registro.distanciaAtual;
      somaDistanciaNova += distanciaNova;
      somaGanho += Math.max(0, registro.distanciaAtual - distanciaNova);
      totalComparavel++;
    });

    regioes.push({
      regiao: grupo.cepBase,
      latitude: arredondarCoordenadaMapa(latitude),
      longitude: arredondarCoordenadaMapa(longitude),
      atendimentos: grupo.atendimentos,
      pessoasDistintas: Object.keys(grupo.pessoas).length,
      peso: grupo.atendimentos,
      distanciaAtualMediaKm: mediaArredondadaMapa(somaDistanciaAtual, totalComparavel),
      distanciaNovaMediaKm: mediaArredondadaMapa(somaDistanciaNova, totalComparavel),
      ganhoMedioKm: mediaArredondadaMapa(somaGanho, totalComparavel),
      ganhoTotalKm: arredondarUmaCasa(somaGanho)
    });
  });

  regioes.sort(function(a, b) {
    if (b.ganhoTotalKm !== a.ganhoTotalKm) return b.ganhoTotalKm - a.ganhoTotalKm;
    if (b.atendimentos !== a.atendimentos) return b.atendimentos - a.atendimentos;
    return normalizar(a.regiao).localeCompare(normalizar(b.regiao));
  });

  const maiorPeso = regioes.reduce(function(maior, item) {
    return Math.max(maior, item.peso || 0);
  }, 1);

  regioes.forEach(function(item) {
    item.intensidade = Math.max(0.18, Math.min(1, item.peso / maiorPeso));
  });

  return regioes;
}

function montarNapsMapaCalor(cargaNaps) {
  return Object.keys(cargaNaps)
    .map(function(chave) {
      const item = cargaNaps[chave];

      return {
        naps: item.naps,
        cepNaps: item.cepNaps,
        latitude: item.latitude,
        longitude: item.longitude,
        atendimentos: item.atendimentos,
        pessoasDistintas: Object.keys(item.pessoas).length,
        distanciaMediaKm: mediaArredondadaMapa(item.somaDistancia, item.totalComDistancia),
        atendimentosDistantes: item.distantes
      };
    })
    .sort(function(a, b) {
      if (b.atendimentos !== a.atendimentos) return b.atendimentos - a.atendimentos;
      if (b.distanciaMediaKm !== a.distanciaMediaKm) return b.distanciaMediaKm - a.distanciaMediaKm;
      return normalizar(a.naps).localeCompare(normalizar(b.naps));
    });
}

function montarResumoMapaCalor(base) {
  return {
    totalAtendimentos: base.totalAtendimentos || 0,
    pessoasDistintas: base.pessoasDistintas || 0,
    regioesMapeadas: base.regioesMapeadas || 0,
    distanciaMediaAtualKm: base.distanciaMediaAtualKm || 0,
    melhorGanhoMedioKm: base.melhorGanhoMedioKm || 0,
    cepsSemCadastro: base.cepsSemCadastro || 0,
    cepsSemCoordenada: base.cepsSemCoordenada || 0,
    registrosSemNaps: base.registrosSemNaps || 0,
    cepsGeocodificadosAgora: base.cepsGeocodificadosAgora || 0,
    limiteGeocodificacaoAtingido: !!base.limiteGeocodificacaoAtingido
  };
}

function montarAvisosMapaCalor(resumo) {
  const avisos = [];

  if (resumo.cepsGeocodificadosAgora > 0) {
    avisos.push("Foram adicionados " + resumo.cepsGeocodificadosAgora + " CEPs ao cache de coordenadas.");
  }

  if (resumo.limiteGeocodificacaoAtingido) {
    avisos.push("Ha muitos CEPs novos. Atualize novamente para completar o cache sem travar a execucao.");
  }

  if (resumo.cepsSemCadastro > 0) {
    avisos.push(resumo.cepsSemCadastro + " atendimento(s) sem CEP de residencia no periodo.");
  }

  if (resumo.cepsSemCoordenada > 0) {
    avisos.push(resumo.cepsSemCoordenada + " atendimento(s) ainda sem coordenada calculada para o CEP.");
  }

  if (resumo.registrosSemNaps > 0) {
    avisos.push(resumo.registrosSemNaps + " atendimento(s) sem CEP do NAPS para calcular distancia atual.");
  }

  return avisos;
}

function mediaArredondadaMapa(soma, total) {
  if (!total) return 0;

  return arredondarUmaCasa(soma / total);
}

function arredondarCoordenadaMapa(valor) {
  return Math.round(Number(valor || 0) * 1000000) / 1000000;
}

function montarPeriodoDashboard(filtros) {
  return {
    dataInicial: filtros.dataInicial,
    dataFinal: filtros.dataFinal,
    rotulo: formatarDataBrasil(filtros.dataInicial) + " a " + formatarDataBrasil(filtros.dataFinal)
  };
}

function montarResumoDashboard(registros, filtros) {
  const resumoRelatorio = montarResumoRelatorio(registros);
  const opms = {};
  const diasComAtendimento = {};
  let atendimentosEmergenciais = 0;
  let atendimentosIndividuais = 0;

  registros.forEach(function(registro) {
    if (ehRegistroFalta(registro)) return;

    const tipo = normalizar(registro.tipoAtendimento);
    const opm = String(registro.opmAtual || "").trim();
    const dataIso = registro.dataCadastroIso || "";

    if (tipo.includes("emergencial")) atendimentosEmergenciais++;
    if (tipo.includes("individual")) atendimentosIndividuais++;
    if (opm) opms[opm] = true;
    if (dataIso) diasComAtendimento[dataIso] = true;
  });

  const diasPeriodo = calcularDiasPeriodoDashboard(filtros);
  const mediaDiaria = diasPeriodo > 0
    ? Math.round((resumoRelatorio.totalAtendimentos / diasPeriodo) * 10) / 10
    : 0;

  return {
    totalAtendimentos: resumoRelatorio.totalAtendimentos,
    pessoasDistintas: resumoRelatorio.pessoasDistintas,
    totalFaltas: resumoRelatorio.totalFaltas,
    totalAltas: resumoRelatorio.totalAltas,
    atendimentosEmergenciais: atendimentosEmergenciais,
    atendimentosIndividuais: atendimentosIndividuais,
    atendimentosFamiliaresOuGrupo: resumoRelatorio.atendimentosFamiliaresOuGrupo,
    atendimentosComVinculos: resumoRelatorio.atendimentosComVinculos,
    totalVinculos: resumoRelatorio.totalVinculos,
    retornos: resumoRelatorio.retornos,
    opmsDistintas: Object.keys(opms).length,
    diasComAtendimento: Object.keys(diasComAtendimento).length,
    mediaDiaria: mediaDiaria
  };
}

function calcularDiasPeriodoDashboard(filtros) {
  const inicio = obterDataInicio(filtros.dataInicial);
  const fim = obterDataFim(filtros.dataFinal);

  if (!inicio || !fim || fim < inicio) return 0;

  return Math.floor((fim.getTime() - inicio.getTime()) / 86400000) + 1;
}

function montarGraficosDashboard(registros, filtros) {
  return {
    porTipo: topDistribuicaoDashboard(contarPorCampo(registros, "tipoAtendimento"), 8),
    porMotivo: topDistribuicaoDashboard(contarPorCampo(registros, "motivo"), 10),
    porOPM: topDistribuicaoDashboard(contarPorCampo(registros, "opmAtual"), 10),
    porNAPS: topDistribuicaoDashboard(contarPorCampo(registros, "naps"), 60),
    comparativoNAPS: montarComparativoNapsDashboard(registros),
    porNAPSAtendimentos: topDistribuicaoDashboard(contarPorNapsDashboard(registros, "atendimentos"), 60),
    porNAPSFaltas: topDistribuicaoDashboard(contarPorNapsDashboard(registros, "faltas"), 60),
    porNAPSAltas: topDistribuicaoDashboard(contarPorNapsDashboard(registros, "altas"), 60),
    porSituacao: topDistribuicaoDashboard(contarPorCampo(registros, "situacaoStatus"), 8),
    porSexo: topDistribuicaoDashboard(contarPorCampo(registros, "sexo"), 8),
    porFaixaEtaria: ordenarFaixaEtariaDashboard(contarPorCampo(registros, "faixaEtaria")),
    evolucao: montarEvolucaoDashboard(registros, filtros)
  };
}

function montarComparativoNapsDashboard(registros) {
  const mapa = {};

  registros.forEach(function(registro) {
    const naps = String(registro.naps || "nao informado").trim() || "nao informado";

    if (!mapa[naps]) {
      mapa[naps] = {
        rotulo: naps,
        atendimentos: 0,
        faltas: 0,
        altas: 0,
        total: 0
      };
    }

    if (ehRegistroFalta(registro)) {
      mapa[naps].faltas++;
    } else {
      mapa[naps].atendimentos++;
    }

    if (ehRegistroAlta(registro)) {
      mapa[naps].altas++;
    }

    mapa[naps].total++;
  });

  return Object.keys(mapa)
    .map(function(chave) {
      return mapa[chave];
    })
    .sort(function(a, b) {
      if (b.total !== a.total) return b.total - a.total;
      return normalizar(a.rotulo).localeCompare(normalizar(b.rotulo));
    });
}

function contarPorNapsDashboard(registros, tipoContagem) {
  const contagem = {};

  registros.forEach(function(registro) {
    const falta = ehRegistroFalta(registro);
    const alta = ehRegistroAlta(registro);

    if (tipoContagem === "atendimentos" && falta) return;
    if (tipoContagem === "faltas" && !falta) return;
    if (tipoContagem === "altas" && !alta) return;

    const naps = String(registro.naps || "nao informado").trim() || "nao informado";
    contagem[naps] = (contagem[naps] || 0) + 1;
  });

  return contagem;
}

function montarRankingsDashboard(registros) {
  return {
    motivos: topDistribuicaoDashboard(contarPorCampo(registros, "motivo"), 10),
    opms: topDistribuicaoDashboard(contarPorCampo(registros, "opmAtual"), 10),
    responsaveis: topDistribuicaoDashboard(contarPorCampo(registros, "responsavel"), 10)
  };
}

function topDistribuicaoDashboard(contagem, limite) {
  return Object.keys(contagem)
    .map(function(chave) {
      return {
        rotulo: chave || "nao informado",
        valor: contagem[chave]
      };
    })
    .sort(function(a, b) {
      if (b.valor !== a.valor) return b.valor - a.valor;
      return normalizar(a.rotulo).localeCompare(normalizar(b.rotulo));
    })
    .slice(0, limite);
}

function ordenarFaixaEtariaDashboard(contagem) {
  const ordem = [
    "ate 17 anos",
    "18 a 29 anos",
    "30 a 39 anos",
    "40 a 49 anos",
    "50 a 59 anos",
    "60 anos ou mais",
    "nao informado"
  ];

  return ordem
    .filter(function(faixa) {
      return contagem[faixa];
    })
    .map(function(faixa) {
      return {
        rotulo: faixa,
        valor: contagem[faixa]
      };
    });
}

function montarEvolucaoDashboard(registros, filtros) {
  const inicio = obterDataInicio(filtros.dataInicial);
  const fim = obterDataFim(filtros.dataFinal);
  const contagem = {};

  registros.forEach(function(registro) {
    if (ehRegistroFalta(registro)) return;

    const data = registro.dataCadastroData;
    if (!data) return;

    const chave = obterMesAno(data);
    contagem[chave] = (contagem[chave] || 0) + 1;
  });

  if (!inicio || !fim || fim < inicio) {
    return {
      titulo: "Evolução mensal",
      agrupamento: "mes",
      pontos: []
    };
  }

  return {
    titulo: "Evolução mensal",
    agrupamento: "mes",
    pontos: montarPontosMensaisDashboard(inicio, fim, contagem)
  };
}

function montarPontosDiariosDashboard(inicio, fim, contagem) {
  const pontos = [];
  const cursor = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate());
  const limite = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate());

  while (cursor <= limite) {
    const chave = formatarDataParaInput(cursor);

    pontos.push({
      rotulo: formatarDiaMesDashboard(cursor),
      chave: chave,
      valor: contagem[chave] || 0
    });

    cursor.setDate(cursor.getDate() + 1);
  }

  return pontos;
}

function montarPontosMensaisDashboard(inicio, fim, contagem) {
  const pontos = [];
  const cursor = new Date(inicio.getFullYear(), inicio.getMonth(), 1);
  const limite = new Date(fim.getFullYear(), fim.getMonth(), 1);

  while (cursor <= limite) {
    const chave = obterMesAno(cursor);

    pontos.push({
      rotulo: formatarMesAnoDashboard(cursor),
      chave: chave,
      valor: contagem[chave] || 0
    });

    cursor.setMonth(cursor.getMonth() + 1);
  }

  return pontos;
}

function formatarDiaMesDashboard(data) {
  return String(data.getDate()).padStart(2, "0") + "/" +
    String(data.getMonth() + 1).padStart(2, "0");
}

function formatarMesAnoDashboard(data) {
  return String(data.getMonth() + 1).padStart(2, "0") + "/" + data.getFullYear();
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

function carregarUsuariosSistemaPorEmail(ss) {
  const sheet = obterOuCriarAba(ss || SpreadsheetApp.getActiveSpreadsheet(), ABA_USUARIOS, CABECALHOS_USUARIOS);
  const dados = sheet.getDataRange().getValues();
  const indices = obterIndicesUsuariosSistema(sheet);
  const mapa = {};

  for (let i = 1; i < dados.length; i++) {
    const usuario = lerDadosUsuarioSistema(dados[i], indices);

    if (usuario.email) {
      mapa[usuario.email] = usuario;
    }
  }

  return mapa;
}

function montarRegistroRelatorio(linha, vinculosPorAtendimento, usuariosPorEmail) {
  const id = linha[0] || "";
  const emailCadastro = String(linha[1] || "").toLowerCase().trim();
  const usuarioCadastro = usuariosPorEmail && usuariosPorEmail[emailCadastro]
    ? usuariosPorEmail[emailCadastro]
    : {};
  const dataCadastroData = obterData(linha[26]);
  const dataNascimentoData = obterData(linha[10]);
  const dataIngressoData = obterData(linha[9]);
  const idade = calcularAnosAteHoje(dataNascimentoData);
  const tempoServicoAnos = calcularAnosAteHoje(dataIngressoData);
  const vinculos = vinculosPorAtendimento[String(id)] || [];

  return {
    id: id,
    emailCadastro: linha[1] || "",
    naps: usuarioCadastro.naps || "nao informado",
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
  let totalFaltas = 0;
  let totalAltas = 0;
  let totalAtendimentos = 0;

  registros.forEach(function(registro) {
    if (ehRegistroFalta(registro)) {
      totalFaltas++;
      return;
    }

    totalAtendimentos++;

    if (ehRegistroAlta(registro)) {
      totalAltas++;
    }

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
    totalRegistros: registros.length,
    totalAtendimentos: totalAtendimentos,
    pessoasDistintas: Object.keys(pessoas).length,
    totalFaltas: totalFaltas,
    totalAltas: totalAltas,
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

function montarDadosIndividuaisRelatorio(filtros, filtrados, todosRegistros, ss) {
  const termo = String(filtros.buscaLivre || "").trim();

  if (!termo) {
    return {
      status: "sem_busca",
      mensagem: "Use a Busca Individual para visualizar estes dados."
    };
  }

  if (!filtrados || filtrados.length === 0) {
    return {
      status: "sem_resultado",
      mensagem: "Nenhum cadastro localizado para a Busca Individual."
    };
  }

  const grupos = agruparRegistrosPorPessoa(filtrados);
  const chaves = Object.keys(grupos);

  if (chaves.length > 1) {
    return {
      status: "multiplos",
      mensagem: "Mais de uma pessoa foi localizada. Refine a busca por CPF completo ou R.E.",
      candidatos: chaves.map(function(chave) {
        const referencia = obterRegistroMaisRecente(grupos[chave]);

        return {
          nome: referencia.nome || "",
          re: referencia.re || "",
          cpf: referencia.cpf || "",
          unidade: referencia.opmAtual || "",
          totalAtendimentos: contarAtendimentosPessoa(todosRegistros, chave)
        };
      }).slice(0, 10)
    };
  }

  const chavePessoa = chaves[0];
  const registrosPessoa = grupos[chavePessoa];
  const referencia = obterRegistroMaisRecente(registrosPessoa);
  const usuarioNaps = obterDadosUsuarioSistemaParaRelatorio(referencia.emailCadastro);
  const distancia = calcularDistanciaResidenciaNaps(referencia.cep, usuarioNaps.cepNaps, ss);

  return {
    status: "ok",
    nome: referencia.nome || "",
    re: referencia.re || "",
    cpf: referencia.cpf || "",
    unidade: referencia.opmAtual || "",
    naps: usuarioNaps.naps || "nao informado",
    totalAtendimentos: contarAtendimentosPessoa(todosRegistros, chavePessoa),
    distanciaKm: distancia.distanciaKm,
    distanciaTexto: distancia.texto,
    distanciaDisponivel: distancia.disponivel,
    dataReferencia: referencia.dataCadastro || "",
    mensagemDistancia: distancia.mensagem || ""
  };
}

function agruparRegistrosPorPessoa(registros) {
  const grupos = {};

  registros.forEach(function(registro) {
    const chave = obterChavePessoa(registro);

    if (!grupos[chave]) grupos[chave] = [];
    grupos[chave].push(registro);
  });

  return grupos;
}

function obterRegistroMaisRecente(registros) {
  return registros.slice().sort(function(a, b) {
    return b.dataCadastroTimestamp - a.dataCadastroTimestamp;
  })[0];
}

function contarAtendimentosPessoa(registros, chavePessoa) {
  let total = 0;

  registros.forEach(function(registro) {
    if (!ehRegistroFalta(registro) && obterChavePessoa(registro) === chavePessoa) {
      total++;
    }
  });

  return total;
}

function ehRegistroFalta(registro) {
  const tipo = normalizar(registro && registro.tipoAtendimento);

  return tipo === "falta" || tipo === "registrar falta" || tipo.includes("no show");
}

function ehRegistroAlta(registro) {
  const tipo = normalizar(registro && registro.tipoAtendimento);

  return tipo === "alta" || tipo === "atendimento de alta";
}

function obterDadosUsuarioSistemaParaRelatorio(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obterOuCriarAba(ss, ABA_USUARIOS, CABECALHOS_USUARIOS);
  const emailNormalizado = String(email || "").toLowerCase().trim();

  if (!emailNormalizado) {
    return {
      email: "",
      nome: "",
      naps: "",
      cepNaps: ""
    };
  }

  const dados = sheet.getDataRange().getValues();
  const indices = obterIndicesUsuariosSistema(sheet);

  for (let i = 1; i < dados.length; i++) {
    const usuarioLinha = lerDadosUsuarioSistema(dados[i], indices);

    if (usuarioLinha.email === emailNormalizado) {
      return {
        email: emailNormalizado,
        nome: usuarioLinha.nome,
        naps: usuarioLinha.naps,
        cepNaps: usuarioLinha.cepNaps
      };
    }
  }

  return {
    email: emailNormalizado,
    nome: "",
    naps: "",
    cepNaps: ""
  };
}

function calcularDistanciaResidenciaNaps(cepResidencia, cepNaps, ss) {
  const cepOrigem = somenteNumeros(cepResidencia);
  const cepDestino = somenteNumeros(cepNaps);

  if (cepOrigem.length !== 8 && cepDestino.length !== 8) {
    return {
      disponivel: false,
      distanciaKm: null,
      texto: "nao disponivel",
      mensagem: "CEP da residencia e CEP do NAPS nao informados."
    };
  }

  if (cepOrigem.length !== 8) {
    return {
      disponivel: false,
      distanciaKm: null,
      texto: "nao disponivel",
      mensagem: "CEP da residencia nao informado."
    };
  }

  if (cepDestino.length !== 8) {
    return {
      disponivel: false,
      distanciaKm: null,
      texto: "nao disponivel",
      mensagem: "CEP do NAPS nao informado."
    };
  }

  if (cepOrigem === cepDestino) {
    return {
      disponivel: true,
      distanciaKm: 0,
      texto: "0 km",
      mensagem: ""
    };
  }

  try {
    const origem = obterCoordenadasPorCEP(cepOrigem, ss);
    const destino = obterCoordenadasPorCEP(cepDestino, ss);

    if (!origem || !destino) {
      return {
        disponivel: false,
        distanciaKm: null,
        texto: "nao disponivel",
        mensagem: "Nao foi possivel localizar coordenadas para um dos CEPs."
      };
    }

    const distanciaKm = arredondarUmaCasa(calcularDistanciaHaversineKm(
      origem.latitude,
      origem.longitude,
      destino.latitude,
      destino.longitude
    ));

    return {
      disponivel: true,
      distanciaKm: distanciaKm,
      texto: String(distanciaKm).replace(".", ",") + " km",
      mensagem: ""
    };
  } catch (erro) {
    return {
      disponivel: false,
      distanciaKm: null,
      texto: "nao disponivel",
      mensagem: erro.message || "Erro ao calcular distancia."
    };
  }
}

function obterCoordenadasPorCEP(cep, ss) {
  const cepNormalizado = somenteNumeros(cep);

  if (cepNormalizado.length !== 8) return null;

  const sheet = obterOuCriarAba(ss || SpreadsheetApp.getActiveSpreadsheet(), ABA_CEPS_CACHE, CABECALHOS_CEPS_CACHE);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const cepCache = somenteNumeros(dados[i][0]);

    if (cepCache === cepNormalizado) {
      const latitude = Number(String(dados[i][1] || "").replace(",", "."));
      const longitude = Number(String(dados[i][2] || "").replace(",", "."));

      if (!isNaN(latitude) && !isNaN(longitude)) {
        return {
          latitude: latitude,
          longitude: longitude
        };
      }
    }
  }

  const endereco = formatarCEP(cepNormalizado) + ", Brasil";
  const resposta = Maps.newGeocoder()
    .setRegion("br")
    .geocode(endereco);

  if (!resposta || resposta.status !== "OK" || !resposta.results || resposta.results.length === 0) {
    return null;
  }

  const localizacao = resposta.results[0].geometry && resposta.results[0].geometry.location;

  if (!localizacao) return null;

  const coordenadas = {
    latitude: Number(localizacao.lat),
    longitude: Number(localizacao.lng)
  };

  if (isNaN(coordenadas.latitude) || isNaN(coordenadas.longitude)) {
    return null;
  }

  sheet.appendRow([
    formatarCEP(cepNormalizado),
    coordenadas.latitude,
    coordenadas.longitude,
    resposta.results[0].formatted_address || endereco,
    new Date()
  ]);

  return coordenadas;
}

function calcularDistanciaHaversineKm(lat1, lon1, lat2, lon2) {
  const raioTerraKm = 6371;
  const dLat = grausParaRadianos(lat2 - lat1);
  const dLon = grausParaRadianos(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(grausParaRadianos(lat1)) *
    Math.cos(grausParaRadianos(lat2)) *
    Math.sin(dLon / 2) *
    Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return raioTerraKm * c;
}

function grausParaRadianos(graus) {
  return graus * Math.PI / 180;
}

function arredondarUmaCasa(valor) {
  return Math.round(Number(valor || 0) * 10) / 10;
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
