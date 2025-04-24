document.getElementById('uploadForm').addEventListener('submit', function(event) {
  event.preventDefault();

  const files = document.getElementById('file').files;
  const tableBody = document.querySelector('#resultTable tbody');

  // Limpar tabela antes de preencher com novos dados
  tableBody.innerHTML = '';

  Array.from(files).forEach(file => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const xmlContent = e.target.result;
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlContent, "text/xml");

      const numeroNF = xmlDoc.getElementsByTagName('nNF')[0]?.textContent || '';
      const emitenteCNPJ = xmlDoc.getElementsByTagName('CNPJ')[0]?.textContent || '';
      const emitenteNome = xmlDoc.getElementsByTagName('xNome')[0]?.textContent || '';
      const dataEmissao = xmlDoc.getElementsByTagName('dhEmi')[0]?.textContent || '';

      const destinatarioCNPJ = xmlDoc.getElementsByTagName('CNPJ')[1]?.textContent || '';
      const destinatarioNome = xmlDoc.getElementsByTagName('xNome')[1]?.textContent || '';

      const produtos = xmlDoc.getElementsByTagName('det');
      Array.from(produtos).forEach(produto => {
        const nomeProduto = produto.getElementsByTagName('xProd')[0]?.textContent || '';
        let quantidade = produto.getElementsByTagName('qCom')[0]?.textContent || '0';
        const valorUnitario = produto.getElementsByTagName('vUnCom')[0]?.textContent || '0';

        quantidade = parseFloat(quantidade.replace(',', '.')) || 0;

        const pis = produto.getElementsByTagName('PIS')[0]?.getElementsByTagName('vPIS')[0]?.textContent || '0';
        const cofins = produto.getElementsByTagName('COFINS')[0]?.getElementsByTagName('vCOFINS')[0]?.textContent || '0';

        const pisValor = parseFloat(pis) || 0;
        const cofinsValor = parseFloat(cofins) || 0;

        const cfop = produto.getElementsByTagName('CFOP')[0]?.textContent || '';
        const ncm = produto.getElementsByTagName('NCM')[0]?.textContent || '';

        const icms = produto.getElementsByTagName('ICMS')[0];
        const cstICMS = icms?.getElementsByTagName('CST')[0]?.textContent || '';

        const cstPIS = produto.getElementsByTagName('PIS')[0]?.getElementsByTagName('CST')[0]?.textContent || '';
        const cstCOFINS = produto.getElementsByTagName('COFINS')[0]?.getElementsByTagName('CST')[0]?.textContent || '';

        const row = document.createElement('tr');
        row.innerHTML = `
          <td>${numeroNF}</td>
          <td>${emitenteCNPJ}</td>
          <td>${emitenteNome}</td>
          <td>${dataEmissao}</td>
          <td>${destinatarioCNPJ}</td>
          <td>${destinatarioNome}</td>
          <td>${nomeProduto}</td>
          <td>${quantidade.toFixed(2)}</td>
          <td>R$ ${parseFloat(valorUnitario).toFixed(2)}</td>
          <td>R$ ${pisValor.toFixed(2)}</td>
          <td>R$ ${cofinsValor.toFixed(2)}</td>
          <td>${cfop}</td>
          <td>${cstICMS}</td>
          <td>${cstPIS}</td>
          <td>${cstCOFINS}</td>
          <td>${ncm}</td>
        `;
        tableBody.appendChild(row);
      });
    };
    reader.readAsText(file);
  });
});

// Botão de download em Excel
const downloadBtn = document.getElementById('downloadExcel');
downloadBtn.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(document.getElementById('resultTable'));
  XLSX.utils.book_append_sheet(wb, ws, 'Notas');
  XLSX.writeFile(wb, 'notas-fiscais.xlsx');
});
