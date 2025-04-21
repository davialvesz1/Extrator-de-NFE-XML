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
  
        const emitenteCNPJ = xmlDoc.getElementsByTagName('CNPJ')[0].textContent;
        const emitenteNome = xmlDoc.getElementsByTagName('xNome')[0].textContent;
        const dataEmissao = xmlDoc.getElementsByTagName('dhEmi')[0].textContent;
  
        const destinatarioCNPJ = xmlDoc.getElementsByTagName('CNPJ')[1]?.textContent || '';
        const destinatarioNome = xmlDoc.getElementsByTagName('xNome')[1]?.textContent || '';
  
        const produtos = xmlDoc.getElementsByTagName('det');
        
        Array.from(produtos).forEach(produto => {
          const nomeProduto = produto.getElementsByTagName('xProd')[0].textContent;
          let quantidade = produto.getElementsByTagName('qCom')[0].textContent;
          const valorUnitario = produto.getElementsByTagName('vUnCom')[0].textContent;
          
          // Verifica se a quantidade é válida, convertendo para número e garantindo que seja no formato adequado
          quantidade = parseFloat(quantidade.replace(',', '.')) || 0; // Substitui vírgula por ponto e tenta converter para número
  
          // PIS e COFINS com verificação de valores vazios
          const pis = produto.getElementsByTagName('PIS')[0]?.getElementsByTagName('vPIS')[0]?.textContent || '0';
          const cofins = produto.getElementsByTagName('COFINS')[0]?.getElementsByTagName('vCOFINS')[0]?.textContent || '0';
  
          // Garantir que valores de PIS e COFINS sejam numéricos
          const pisValor = parseFloat(pis) || 0;
          const cofinsValor = parseFloat(cofins) || 0;
  
          const row = document.createElement('tr');
          row.innerHTML = `
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
          `;
          tableBody.appendChild(row);
        });
      };
      reader.readAsText(file);
    });
  });
  document.getElementById("downloadExcel").addEventListener("click", function () {
    const table = document.getElementById("resultTable");
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(table);
    XLSX.utils.book_append_sheet(wb, ws, "Notas Fiscais");
  
    const now = new Date();
    const timestamp = now.toISOString().slice(0, 19).replace(/[-:T]/g, "");
    const filename = `notas_fiscais_${timestamp}.xlsx`;
  
    XLSX.writeFile(wb, filename);
  });
  