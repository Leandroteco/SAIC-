const ABA_DADOS = "dados_cadastro";
const ABA_VINCULOS = "pessoas_vinculadas";

function doGet(e) {
  if (e && e.parameter && e.parameter.pagina === "relatorios") {
    return HtmlService.createHtmlOutputFromFile("Relatorios")
      .setTitle("SAIC - Relatórios Gerenciais");
  }

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("SAIC - Atendimento Integrado dos NAPS");
}

function salvarAtendimento(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_DADOS);
  const sheetVinculos = ss.getSheetByName(ABA_VINCULOS);

  if (!sheet) throw new Error("A aba dados_cadastro não foi encontrada.");
  if (!sheetVinculos) throw new Error("A aba pessoas_vinculadas não foi encontrada.");

  const idAtendimento = gerarNovoId(sheet, 1);
  const dataCadastro = new Date();

 sheet.appendRow([
  idAtendimento,
  normalizar(dados.naps),
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
  dataCadastro
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
function buscarCadastro(termo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados_cadastro");
  const linhas = sheet.getDataRange().getValues();

  const termoOriginal = String(termo || "").trim();
  const buscaTexto = normalizar(termoOriginal);
  const buscaNumeros = somenteNumeros(termoOriginal);

  let resultados = [];

  const contemLetras = /[A-Z]/i.test(termoOriginal);
  const pareceNome = contemLetras && !/^\d{1,6}-?[\dA]?$/i.test(termoOriginal);
  const pareceCPF = buscaNumeros.length >= 8;
  const pareceRE = !pareceNome && buscaNumeros.length > 0 && buscaNumeros.length <= 7;

  for (let i = linhas.length - 1; i >= 1; i--) {
    const l = linhas[i];

    const registro = {
      re: l[4] || "",
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
      complemento: l[23] || ""
    };

    const reTexto = normalizar(registro.re);
    const reNumerico = somenteNumeros(registro.re);
    const nomeTexto = normalizar(registro.nome);
    const cpfNumerico = somenteNumeros(registro.cpf);

    let achou = false;

    if (pareceCPF) {
      achou = cpfNumerico.includes(buscaNumeros);
    } else if (pareceRE) {
     achou = reNumerico.startsWith(buscaNumeros) || reTexto.includes(buscaTexto);
    } else {
      achou = nomeTexto.includes(buscaTexto);
    }

    if (achou) {
      resultados.push(registro);
    }
  }

  if (resultados.length === 0) {
    return {
      encontrado: false,
      mensagem: "Nenhum cadastro anterior localizado. Preencha novo cadastro."
    };
  }

  // Se for CPF completo ou RE completo, preenche direto com o registro mais recente
  const reCompleto = /^[0-9]{6}-[0-9A]$/i.test(termoOriginal);
const cpfCompleto = buscaNumeros.length === 11;

if ((reCompleto || cpfCompleto) && resultados.length >= 1) {
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

  // Em todos os demais casos, mostra lista de opções
  return {
    encontrado: true,
    multiplos: true,
    resultados: resultados.slice(0, 10)
  };
}

function montarRegistro(linha) {
  return {
    id: linha[0],
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
    dataCadastro: linha[26]
  };
}

function formatarDataParaInput(valor) {
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    const ano = valor.getFullYear();
    const mes = String(valor.getMonth() + 1).padStart(2, "0");
    const dia = String(valor.getDate()).padStart(2, "0");
    return `${ano}-${mes}-${dia}`;
  }

  return "";
}

// BUSCA POR RE, CPF E NOME


function buscarCadastro(termo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados_cadastro");
  const linhas = sheet.getDataRange().getValues();

  const buscaTexto = normalizar(termo);
  const buscaNumeros = somenteNumeros(termo);

  let resultados = [];
  let chavesEncontradas = {};

  for (let i = linhas.length - 1; i >= 1; i--) {
    const l = linhas[i];

    const registro = {
      re: l[4] || "",
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
      complemento: l[23] || ""
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

  // CPF ou RE completo: preenche direto
  if (buscaNumeros.length >= 11 || /^[0-9]{6}-[0-9A]$/i.test(String(termo).trim())) {
    return {
      encontrado: true,
      multiplos: false,
      registro: resultados[0]
    };
  }

  // Se encontrou apenas um cadastro pelo nome, preenche direto.
// Se encontrou mais de um, mostra lista.
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
  if (!valor) return "";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    const ano = valor.getFullYear();
    const mes = String(valor.getMonth() + 1).padStart(2, "0");
    const dia = String(valor.getDate()).padStart(2, "0");
    return `${ano}-${mes}-${dia}`;
  }

  return "";
}
function gerarRelatorioGerencial() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dados_cadastro");
  const dados = sheet.getDataRange().getValues();

  const totalAtendimentos = dados.length - 1;

  const porNAPS = {};
  const porTipo = {};
  const porOPM = {};
  const porCidade = {};
  const porResponsavel = {};

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    const naps = linha[1] || "nao informado";
    const tipoAtendimento = linha[2] || "nao informado";
    const opm = linha[12] || "nao informado";
    const cidade = linha[20] || "nao informado";
    const responsavel = linha[25] || "nao informado";

    porNAPS[naps] = (porNAPS[naps] || 0) + 1;
    porTipo[tipoAtendimento] = (porTipo[tipoAtendimento] || 0) + 1;
    porOPM[opm] = (porOPM[opm] || 0) + 1;
    porCidade[cidade] = (porCidade[cidade] || 0) + 1;
    porResponsavel[responsavel] = (porResponsavel[responsavel] || 0) + 1;
  }

  return {
    totalAtendimentos: totalAtendimentos,
    porNAPS: porNAPS,
    porTipo: porTipo,
    porOPM: porOPM,
    porCidade: porCidade,
    porResponsavel: porResponsavel
  };
}

function obterUrlApp() {
  return ScriptApp.getService().getUrl();
}
