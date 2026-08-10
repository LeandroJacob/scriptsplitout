const XLSX = require('xlsx');

exports.handler = async (event) => {
  try {
    if (event.httpMethod !== 'POST') {
      return { statusCode: 405, body: JSON.stringify({ error: 'Método não permitido. Use POST.' }) };
    }
    if (!event.body) {
      return { statusCode: 400, body: JSON.stringify({ error: 'Corpo da requisição vazio.' }) };
    }
    const bodyBuffer = Buffer.from(event.body, 'base64');
    const workbook = XLSX.read(bodyBuffer, { type: 'buffer' });

    const getHeaders = (sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || !sheet['!ref']) return [];
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const headers = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
        headers.push(cell ? cell.v : null);
      }
      return headers;
    };

    // NOVO: extrai a linha 1 (primeira linha de dados, após o header) como referência de FORMATO.
    // Usado só para inferir tipo/padrão de valor (ex: numérico vs texto, "N"/"S", código vs texto livre).
    // Quem chama esta function NUNCA deve reencaminhar isto para o array de saída da carga —
    // é dado de diagnóstico, não deve ser persistido nem reenviado à Meweb.
    const getSampleRow = (sheetName, headers) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || !sheet['!ref']) return {};
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const row = {};
      for (let col = range.s.c; col <= range.e.c; col++) {
        const header = headers[col - range.s.c];
        if (!header) continue;
        const cell = sheet[XLSX.utils.encode_cell({ r: 1, c: col })]; // linha 1 = primeira linha de dados
        row[header] = cell ? cell.v : null;
      }
      return row;
    };

    const includeSample = event.queryStringParameters?.include_sample === 'true';
    const mantisColumns = getHeaders('Mantis');
    const itensColumns = getHeaders('ItensRequisicao');

    const response = {
      sheet_names: workbook.SheetNames,
      mantis_columns: mantisColumns,
      itens_columns: itensColumns,
    };

    if (includeSample) {
      response.mantis_sample_row = getSampleRow('Mantis', mantisColumns);
      response.itens_sample_row = getSampleRow('ItensRequisicao', itensColumns);
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(response),
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message, stack: err.stack }) };
  }
};
