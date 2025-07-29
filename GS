function formatarDataHora(data) {
  if (!data) return '';
  const d = new Date(data);
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const ano = d.getFullYear();
  const hora = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${dia}/${mes}/${ano} ${hora}:${min}`;
}

function formatarData(data) {
  if (!data) return '';
  const d = new Date(data);
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const ano = d.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

function parseDataISO(dataStr) {
  if (!dataStr) return null;

  const isoRegex = /^\d{4}-\d{2}-\d{2}$/;
  if (isoRegex.test(dataStr)) {
    const [ano, mes, dia] = dataStr.split("-");
    return new Date(ano, mes - 1, dia);
  }

  const brRegex = /^\d{2}\/\d{2}\/\d{4}$/;
  if (brRegex.test(dataStr)) {
    const [dia, mes, ano] = dataStr.split("/");
    return new Date(ano, mes - 1, dia);
  }

  return new Date(dataStr);
}

function SalvarEvento(dados) {
  try {
    const ID_PLANILHA_RIO = "1OftceaPoeqJBSwRGyx3XNcYV_3hzHc2J-7xNauEck6Q";
    const ID_PLANILHA_ZIPDIN = "1qeLYnHidP4n10DIi-iZbOwHs0vuaLxBl6oV86Haqf6g";

    const planilhaId = (dados['empresa'] === 'rio') ? ID_PLANILHA_RIO : ID_PLANILHA_ZIPDIN;
    const ss = SpreadsheetApp.openById(planilhaId);
    const nomeEmpresa = (dados['empresa'] === 'rio') ? 'RioCred' : 'ZipDin';

    const dataHoraRegistro = new Date();

    if (dados['vencimento']) {
      const dataVenc = parseDataISO(dados['vencimento']);
      const ano = dataVenc.getFullYear();
      if (ano < 1900 || ano > 2032) {
        return { sucesso: false, mensagem: "Data de vencimento inválida. Aceito apenas datas entre 1900 e 2032." };
      }
    }

    if (dados['data_pagamento']) {
      const dataPgto = parseDataISO(dados['data_pagamento']);
      const ano = dataPgto.getFullYear();
      if (ano < 1900 || ano > 2032) {
        return { sucesso: false, mensagem: "Data de pagamento inválida. Aceito apenas datas entre 1900 e 2032." };
      }
    }

    if (dados['vencimento_boleto']) {
      const dataBoleto = parseDataISO(dados['vencimento_boleto']);
      const ano = dataBoleto.getFullYear();
      if (ano < 1900 || ano > 2032) {
        return { sucesso: false, mensagem: "Vencimento do boleto inválido. Aceito apenas datas entre 1900 e 2032." };
      }
    }

    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    const nomeAba = `${meses[dataHoraRegistro.getMonth()]}-${dataHoraRegistro.getFullYear()}`;

    let aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      aba = ss.insertSheet(nomeAba);
      aba.getRange('A1:P1').setValues([[
        'ID', 'Data e Hora do Registro', 'Operador', 'Cliente', 'CPF do Cliente',
        'Contrato', 'Parcela', 'Vencimento Original da Parcela', 'Vencimento do Boleto',
        'Data Pagamento', 'Atraso', 'Evento', 'Telefone', 'Canal', 'Status', 'Observações'
      ]]);
    }

    const ultimaLinha = aba.getLastRow();
    const linhaInserir = (ultimaLinha < 4) ? 4 : ultimaLinha + 1;
    const id = linhaInserir - 3;

    let dataPagamento = '';
    let atraso = '';

    const dtPgto = parseDataISO(dados['data_pagamento']);
    const dtVenc = parseDataISO(dados['vencimento']);

    if (dtPgto instanceof Date && !isNaN(dtPgto)) {
      dataPagamento = formatarData(dtPgto);
    }

    if (dtPgto instanceof Date && !isNaN(dtPgto) && dtVenc instanceof Date && !isNaN(dtVenc)) {
      const difDias = Math.ceil((dtPgto - dtVenc) / (1000 * 60 * 60 * 24));
      atraso = difDias.toString();
    }

    const vencOriginal = dados['vencimento'] ? formatarData(parseDataISO(dados['vencimento'])) : '';
    const vencBoleto = dados['vencimento_boleto'] ? formatarData(parseDataISO(dados['vencimento_boleto'])) : '';

    const linhaDados = [
      id,
      formatarDataHora(dataHoraRegistro),
      dados['operador'] || '',
      dados['cliente'] || '',
      dados['cpf'] || '',
      dados['contrato'] || '',
      dados['parcela'] || '',
      vencOriginal,
      vencBoleto,
      dataPagamento,
      atraso,
      dados['evento'] || '',
      dados['telefone'] || '',
      dados['canal'] || '',
      dados['status'] || '',
      dados['observacoes'] || ''
    ];

    aba.getRange(linhaInserir, 1, 1, linhaDados.length).setValues([linhaDados]);

    return { sucesso: true, mensagem: `Registro salvo com sucesso! Empresa: ${nomeEmpresa} | Aba: ${nomeAba}` };

  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao salvar evento: " + e.message };
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('PaginaEvento')
    .setTitle("Registro de Eventos")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
