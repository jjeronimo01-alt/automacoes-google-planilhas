/**
 * =========================================================================
 * DASHBOARD GESTÃO - MÓDULO CENTRAL DE UTILITÁRIOS (Utils.gs)
 * Funções compartilhadas entre Lanc_Diario.gs, Lanc_Semanal.gs e Lanc_Mensal.gs
 * =========================================================================
 */

/**
 * Converte string 'YYYY-MM-DD' em objeto Date local (meia-noite)
 */
function parseDataIso(strIso) {
  const [y, m, d] = String(strIso).split('-').map(Number);
  return new Date(y, m - 1, d);
}

/**
 * Converte string 'YYYY-MM' em objeto Date do 1º dia do mês
 */
function parseMesIso(mesIso) {
  const [y, m] = String(mesIso).split('-').map(Number);
  return new Date(y, m - 1, 1);
}

/**
 * Normaliza qualquer entrada de data (Date, ISO UTC com 'T', DD/MM/YYYY, YYYY-MM-DD)
 * para chave 'YYYY-MM-DD' respeitando o fuso horário de Brasília (GMT-0300).
 */
function normalizarDataChave(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, "GMT-0300", 'yyyy-MM-dd');
  }

  const rawStr = String(valor).trim();
  if (!rawStr) return null;

  // Tratamento de Timestamps ISO/UTC (ex: '2026-09-28T01:15:00+00:00' ou 'Z' da Amazon/ML)
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(rawStr)) {
    const dtParsed = new Date(rawStr);
    if (!isNaN(dtParsed.getTime())) {
      return Utilities.formatDate(dtParsed, "GMT-0300", 'yyyy-MM-dd');
    }
  }

  const str = rawStr.split(' ')[0].split('T')[0];

  // Formato DD/MM/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    const [d, m, y] = str.split('/');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  // Formato DD/MM/YY
  if (/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(str)) {
    const [d, m, y] = str.split('/');
    return `20${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  // Formato YYYY-MM-DD
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
    const [y, m, d] = str.split('-');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  return null;
}

/**
 * Converte datas em extenso da Amazon (ex: '15 de set. de 2026') em Date
 */
function parseDataAmazonExtenso(textoData) {
  if (!textoData) return null;
  if (textoData instanceof Date) return textoData;

  const mesesMap = {
    'jan': 0, 'fev': 1, 'mar': 2, 'abr': 3, 'mai': 4, 'jun': 5,
    'jul': 6, 'ago': 7, 'set': 8, 'out': 9, 'nov': 10, 'dez': 11
  };
  const limpo = String(textoData).toLowerCase().replace(/de/g, ' ').replace(/\./g, ' ').trim();
  const partes = limpo.split(/\s+/);

  if (partes.length >= 3) {
    const dia = parseInt(partes[0], 10);
    const mesStr = partes[1].substring(0, 3);
    const ano = parseInt(partes[2], 10);

    if (!isNaN(dia) && mesesMap[mesStr] !== undefined && !isNaN(ano)) {
      return new Date(ano, mesesMap[mesStr], dia);
    }
  }

  const isoStr = normalizarDataChave(textoData);
  return isoStr ? parseDataIso(isoStr) : null;
}

/**
 * Converte valores monetários ou textuais em Number positivo/padrão
 */
function limparNumero(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let str = String(val).replace('R$', '').replace(/\s/g, '').trim();
  if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

/**
 * Converte valores monetários preservando sinal negativo (ex: '-R$ 150,00' ou '(150,00)')
 */
function limparNumeroComSinal(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let str = String(val).replace('R$', '').replace(/\s/g, '').trim();
  const isNegativo = str.includes('-') || (str.startsWith('(') && str.endsWith(')'));
  str = str.replace(/[()\-]/g, '');

  if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }

  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return isNegativo ? -num : num;
}

/**
 * Formata 'YYYY-MM-DD' para 'DD/MM/YYYY'
 */
function formatarDataBR(dataChave) {
  if (!dataChave) return '';
  const [y, m, d] = String(dataChave).split('-');
  return `${d}/${m}/${y}`;
}

/**
 * Retorna o nome do dia da semana em português
 */
function obterNomeDiaSemana(dataStr) {
  const [y, m, d] = String(dataStr).split('-').map(Number);
  const data = new Date(y, m - 1, d);
  const dias = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];
  return dias[data.getDay()];
}

/**
 * Formata 'YYYY-MM' em rótulo 'set./2026'
 */
function formatarMesRotulo(mesIso) {
  const [y, m] = String(mesIso).split('-');
  const mesesAbrev = ['jan.', 'fev.', 'mar.', 'abr.', 'mai.', 'jun.', 'jul.', 'ago.', 'set.', 'out.', 'nov.', 'dez.'];
  return `${mesesAbrev[parseInt(m, 10) - 1]}/${y}`;
}

/**
 * Converte rótulo ('set./2026'), string ISO ('2026-09-01') ou Date em 'YYYY-MM'
 */
function converterRotuloParaIso(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, "GMT-0300", 'yyyy-MM');
  }
  const str = String(valor).trim();
  if (/^\d{4}-\d{2}/.test(str)) {
    return str.substring(0, 7);
  }
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    const [, m, y] = str.split('/');
    return `${y}-${m.padStart(2, '0')}`;
  }
  const mesesMap = {
    'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06',
    'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12'
  };
  const limpo = str.replace('.', '').replace('/', ' ').trim();
  const partes = limpo.split(/\s+/);
  if (partes.length === 2) {
    const mStr = partes[0].substring(0, 3).toLowerCase();
    const yStr = partes[1].length === 2 ? `20${partes[1]}` : partes[1];
    if (mesesMap[mStr]) return `${yStr}-${mesesMap[mStr]}`;
  }
  return null;
}

/**
 * Obtém a segunda-feira ISO ('YYYY-MM-DD') correspondente a uma data
 */
function obterSegundaFeiraIso(dataStr) {
  const [y, m, d] = String(dataStr).split('-').map(Number);
  const data = new Date(y, m - 1, d);
  const diaSemana = data.getDay();
  const distanciaSegunda = (diaSemana + 6) % 7;
  data.setDate(data.getDate() - distanciaSegunda);
  return Utilities.formatDate(data, "GMT-0300", 'yyyy-MM-dd');
}

/**
 * Calcula o número da semana ISO (1 a 53)
 */
function obterNumeroSemanaIso(d) {
  const data = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const diaNum = data.getUTCDay() || 7;
  data.setUTCDate(data.getUTCDate() + 4 - diaNum);
  const anoInicio = new Date(Date.UTC(data.getUTCFullYear(), 0, 1));
  return Math.ceil((((data - anoInicio) / 86400000) + 1) / 7);
}

/**
 * Localiza a linha de cabeçalho que contenha os termos-chave informados
 */
function localizarIndiceCabecalho(matriz, termosChave) {
  for (let r = 0; r < Math.min(matriz.length, 30); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    for (let t = 0; t < termosChave.length; t++) {
      if (linhaStr.includes(termosChave[t])) {
        return r;
      }
    }
  }
  return 0;
}

/**
 * Leitor unificado e blindado para Google Sheets, XLSX, XLS, CSV, TXT e TSV.
 * Suporta seleção de aba específica (ex: 'Relatório de campanha' do ML Ads).
 */
function lerDadosArquivoUnificado(arquivo, nomeAba) {
  if (!arquivo) return null;
  const mimeType = arquivo.getMimeType();
  const nome = arquivo.getName().toLowerCase();

  Logger.log(`Lendo arquivo: "${arquivo.getName()}" (${mimeType})`);

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    const ss = SpreadsheetApp.openById(arquivo.getId());
    let sheet = null;
    if (nomeAba) {
      sheet = ss.getSheetByName(nomeAba) || ss.getSheets().find(s => s.getName().toLowerCase().includes(nomeAba.toLowerCase()));
    }
    if (!sheet) sheet = ss.getSheets()[0];
    return sheet.getDataRange().getValues();
  }

  if (nome.endsWith('.xlsx') || nome.endsWith('.xls') || mimeType.includes('spreadsheetml') || mimeType.includes('ms-excel')) {
    let tempFileId = null;
    try {
      const blob = arquivo.getBlob();
      const recurso = {
        title: "temp_convert_" + new Date().getTime(),
        mimeType: MimeType.GOOGLE_SHEETS
      };
      const fileInserido = Drive.Files.insert(recurso, blob, { convert: true });
      tempFileId = fileInserido.id;

      const ssConvertida = SpreadsheetApp.openById(tempFileId);
      const todasAbas = ssConvertida.getSheets();
      let abaAlvo = null;

      if (nomeAba) {
        const nomeNorm = nomeAba.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
        abaAlvo = todasAbas.find(s => {
          const sNorm = s.getName().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
          return sNorm.includes("campanha") || sNorm.includes("relat") || sNorm === nomeNorm;
        });
        if (!abaAlvo && todasAbas.length >= 3) {
          abaAlvo = todasAbas[2];
        }
      }

      if (!abaAlvo) {
        abaAlvo = todasAbas[0];
      }

      return abaAlvo.getDataRange().getValues();
    } catch (e) {
      Logger.log("Erro na conversão automática do Excel: " + e.message);
    } finally {
      if (tempFileId) {
        try { Drive.Files.remove(tempFileId); } catch (ex) {}
      }
    }
  }

  // Arquivos de texto (CSV, TXT, TSV)
  let texto = '';
  try {
    texto = arquivo.getBlob().getDataAsString('UTF-8');
    if (texto.includes('\uFFFD')) {
      texto = arquivo.getBlob().getDataAsString('ISO-8859-1');
    }
  } catch (e) {
    texto = arquivo.getBlob().getDataAsString('ISO-8859-1');
  }

  const linhas = texto.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (linhas.length === 0) return [];

  let maxVirgula = 0, maxPontoVirgula = 0, maxTab = 0;
  for (let r = 0; r < Math.min(linhas.length, 15); r++) {
    maxVirgula = Math.max(maxVirgula, (linhas[r].match(/,/g) || []).length);
    maxPontoVirgula = Math.max(maxPontoVirgula, (linhas[r].match(/;/g) || []).length);
    maxTab = Math.max(maxTab, (linhas[r].match(/\t/g) || []).length);
  }

  let delim = ',';
  if (maxTab > maxVirgula && maxTab > maxPontoVirgula) delim = '\t';
  else if (maxPontoVirgula >= maxVirgula) delim = ';';

  return linhas.map(linha => {
    const valores = [];
    let dentroAspas = false;
    let atual = '';
    for (let i = 0; i < linha.length; i++) {
      const c = linha[i];
      if (c === '"') dentroAspas = !dentroAspas;
      else if (c === delim && !dentroAspas) {
        valores.push(atual.trim());
        atual = '';
      } else {
        atual += c;
      }
    }
    valores.push(atual.trim());
    return valores;
  });
}

/**
 * Busca recursivamente o arquivo mais recente dentro de uma pasta
 */
function obterArquivoMaisRecenteRecursivo(pasta) {
  if (!pasta) return null;
  let maisRecente = null;
  let maiorTime = 0;

  function vasculhar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const arq = files.next();
      const nome = arq.getName();
      if (nome.startsWith('~') || nome.startsWith('.')) continue;
      const t = arq.getLastUpdated().getTime();
      if (t > maiorTime) {
        maiorTime = t;
        maisRecente = arq;
      }
    }
    const sub = f.getFolders();
    while (sub.hasNext()) vasculhar(sub.next());
  }

  vasculhar(pasta);
  return maisRecente;
}

/**
 * Retorna todos os arquivos válidos dentro de uma pasta e subpastas
 */
function obterTodosArquivosRecursivo(pasta) {
  if (!pasta) return [];
  const lista = [];
  function vasculhar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const arq = files.next();
      const n = arq.getName();
      if (!n.startsWith('~') && !n.startsWith('.')) lista.push(arq);
    }
    const sub = f.getFolders();
    while (sub.hasNext()) vasculhar(sub.next());
  }
  vasculhar(pasta);
  return lista;
}

/**
 * Localiza a subpasta referente ao mês 'YYYY-MM'.
 * TRAVA DE SEGURANÇA: Se a pasta raiz estiver organizada em subpastas mensais
 * e a subpasta do mês alvo ainda não existir, retorna null (evitando puxar arquivos do mês anterior).
 */
function obterSubpastaDoMes(idPastaRaiz, mesIso) {
  const pastaRaiz = DriveApp.getFolderById(idPastaRaiz);
  const [ano, mesNum] = String(mesIso).split('-');
  const mesesNomes = [
    'janeiro', 'fevereiro', 'março', 'marco', 'abril', 'maio', 'junho',
    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
  ];
  const nomeMesAlvo = mesesNomes[parseInt(mesNum, 10) - 1];

  const subpastas = pastaRaiz.getFolders();
  let possuiSubpastasDeMeses = false;

  while (subpastas.hasNext()) {
    const sub = subpastas.next();
    const nomeSub = sub.getName().toLowerCase();

    if (/\d{4}[-/]\d{2}|\b(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b/i.test(nomeSub)) {
      possuiSubpastasDeMeses = true;
    }

    if (
      nomeSub.includes(mesIso) ||
      (nomeSub.includes(nomeMesAlvo) && nomeSub.includes(ano)) ||
      nomeSub.includes(`${mesNum}/${ano}`) ||
      nomeSub.includes(`${mesNum}-${ano}`)
    ) {
      return sub;
    }
  }

  return possuiSubpastasDeMeses ? null : pastaRaiz;
}

/**
 * Obtém o arquivo consolidado ou mais recente de uma pasta mensal específica
 */
function obterArquivoMaisRecenteOuConsolidado(pasta) {
  if (!pasta) return null;
  let melhorArquivo = null;
  let maiorTime = 0;

  function vasculhar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const arq = files.next();
      const nome = arq.getName().toLowerCase();
      if (nome.startsWith('~') || nome.startsWith('.')) continue;

      if (nome.includes('consolidado') || nome.includes('mensal')) return arq;

      const t = arq.getLastUpdated().getTime();
      if (t > maiorTime) {
        maiorTime = t;
        melhorArquivo = arq;
      }
    }
    const subs = f.getFolders();
    while (subs.hasNext()) {
      const arqConsol = vasculhar(subs.next());
      if (arqConsol) return arqConsol;
    }
    return null;
  }

  const arquivoConsolidado = vasculhar(pasta);
  return arquivoConsolidado || melhorArquivo;
}