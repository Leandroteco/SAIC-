const ABA_DADOS = "dados_cadastro";
const ABA_VINCULOS = "pessoas_vinculadas";

function doGet(e) {
  if (e && e.parameter && e.parameter.pagina === "relatorios") {
    return HtmlService.createHtmlOutputFromFile("relatorios")
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
  const data = obterData(valor);

  if (!data) return "";

  const ano = data.getFullYear();
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  const dia = String(data.getDate()).padStart(2, "0");

  return `${ano}-${mes}-${dia}`;
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
  return formatarDataParaInput(valor);
}

function gerarRelatorioGerencial(filtrosOuDataInicial, dataFinal, tiposRelatorio) {
  const filtros = normalizarFiltrosRelatorio(filtrosOuDataInicial, dataFinal, tiposRelatorio);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_DADOS);

  if (!sheet) throw new Error("A aba dados_cadastro não foi encontrada.");

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
    tipoVinculo: normalizar(filtros.tipoVinculo),
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
      registro.re,
      registro.cpf,
      registro.tipoAtendimento,
      registro.motivo,
      registro.naps,
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

  if (filtros.tipoVinculo && !registro.vinculos.some(function(vinculo) {
    return normalizar(vinculo.tipoVinculo) === filtros.tipoVinculo;
  })) {
    return false;
  }

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
    naps: linha[1] || "",
    tipoAtendimento: linha[2] || "",
    motivo: linha[3] || "",
    re: linha[4] || "",
    nome: linha[5] || "",
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
    porTipoVinculo: contarPorVinculo(registros, "tipoVinculo"),
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
        nome: registro.nome,
        re: registro.re,
        cpf: registro.cpf,
        telefone: registro.telefone,
        tipoAtendimento: registro.tipoAtendimento,
        motivo: registro.motivo,
        naps: registro.naps,
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
    tipoVinculo: listarUnicosRelatorio(todosVinculos, "tipoVinculo"),
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
    `id-${registro.id}`;
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

  return `${ano}-${mes}`;
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

  return `${dia}/${mes}/${ano}`;
}
