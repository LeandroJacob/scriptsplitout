const XLSX = require('xlsx');

exports.handler = async (event) => {
  try {
    if (event.httpMethod !== 'POST') {
      return {
        statusCode: 405,
        body: JSON.stringify({ error: 'Método não permitido. Use POST.' }),
      };
    }

    const bodyBuffer = event.isBase64Encoded
      ? Buffer.from(event.body, 'base64')
      : Buffer.from(event.body, 'binary');

    const workbook = XLSX.read(bodyBuffer, { type: 'buffer' });

    const getHeaders = (sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return [];
      if (!sheet['!ref']) return [];

      const range = XLSX.utils.decode_range(sheet['!ref']);
      const headers = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        const cell = sheet[cellAddress];
        headers.push(cell ? cell.v : null);
      }
      return headers;
    };

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sheet_names: workbook.SheetNames,
        mantis_columns: getHeaders('Mantis'),
        itens_columns: getHeaders('ItensRequisicao'),
      }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
