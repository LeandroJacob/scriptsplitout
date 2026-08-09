const XLSX = require('xlsx');

exports.handler = async function (event) {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method not allowed. Use POST.' })
    };
  }

  try {
    console.log('Content-Type recebido:', event.headers && event.headers['content-type']);
    console.log('event.body (primeiros 500 caracteres):', (event.body || '').slice(0, 500));
    console.log('typeof event.body:', typeof event.body);
    console.log('isBase64Encoded:', event.isBase64Encoded);

    const rawBody = event.isBase64Encoded
      ? Buffer.from(event.body || '', 'base64').toString('utf8')
      : (event.body || '{}');

    const payload = JSON.parse(rawBody);
    const { mantis, itensRequisicao } = payload;

    console.log('mantis is array:', Array.isArray(mantis), 'length:', mantis && mantis.length);
    console.log('itensRequisicao is array:', Array.isArray(itensRequisicao), 'length:', itensRequisicao && itensRequisicao.length);

    if (!Array.isArray(mantis) || !Array.isArray(itensRequisicao)) {
      return {
        statusCode: 400,
        body: JSON.stringify({
          error: "Body invalido. Esperado: { mantis: [...], itensRequisicao: [...] }"
        })
      };
    }

    if (mantis.length === 0 || itensRequisicao.length === 0) {
      return {
        statusCode: 400,
        body: JSON.stringify({
          error: "mantis e itensRequisicao nao podem estar vazios"
        })
      };
    }

    const wb = XLSX.utils.book_new();

    const wsMantis = XLSX.utils.json_to_sheet(mantis);
    const wsItens = XLSX.utils.json_to_sheet(itensRequisicao);

    XLSX.utils.book_append_sheet(wb, wsMantis, 'Mantis');
    XLSX.utils.book_append_sheet(wb, wsItens, 'ItensRequisicao');

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="carga.xlsx"'
      },
      body: buffer.toString('base64'),
      isBase64Encoded: true
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
