const ABA_DADOS = "dados_cadastro";
const ABA_VINCULOS = "pessoas_vinculadas";

function doGet() {
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_DADOS);

  if (!sheet) {
    throw new Error("A aba dados_cadastro não foi encontrada.");
  }

  const dados = sheet.getDataRange().getValues();

  if (dados.length < 2) {
    return {
      encontrado: false,
      mensagem: "Nenhum cadastro anterior localizado."
    };
  }

  const termoOriginal = String(termo || "").trim();
  const termoNormalizado = normalizar(termoOriginal);
  const termoNumerico = somenteNumeros(termoOriginal);

  let resultados = [];

  for (let i = dados.length - 1; i >= 1; i--) {
    const linha = dados[i];

    const registro = montarRegistro(linha);

    const cpfNumerico = somenteNumeros(registro.cpf);
    const reNormalizado = normalizar(registro.re);
    const nomeNormalizado = normalizar(registro.nome);

    const encontrouCPF =
      termoNumerico.length >= 3 && cpfNumerico.includes(termoNumerico);

    const encontrouRE =
      termoNormalizado.length >= 2 && reNormalizado.includes(termoNormalizado);

    const encontrouNome =
      termoNormalizado.length >= 3 && nomeNormalizado.includes(termoNormalizado);

    if (encontrouCPF || encontrouRE || encontrouNome) {
      resultados.push(registro);
    }
  }

  if (resultados.length === 0) {
    return {
      encontrado: false,
      mensagem: "Nenhum cadastro anterior localizado. Preencha novo cadastro."
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
