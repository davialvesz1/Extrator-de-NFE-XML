document.getElementById('uploadForm').addEventListener('submit', async function(event) {
  event.preventDefault();

  // Mostrar mensagem de carregamento
  document.getElementById('loadingMessage').style.display = 'block';

  const files = document.getElementById('file').files;
  const tableBody = document.querySelector('#resultTable tbody');
  tableBody.innerHTML = '';

  if (!files.length) {
    alert("Nenhum arquivo selecionado.");
    // Esconder mensagem caso erro
    document.getElementById('loadingMessage').style.display = 'none';
    return;
  }

  // Remove namespaces de XMLs
  const removeNamespaces = (xmlDoc) => {
    const removeNS = (node) => {
      if (node.prefix) node.nodeName = node.localName;
      if (node.attributes) {
        for (let i = node.attributes.length - 1; i >= 0; i--) {
          const attr = node.attributes[i];
          if (attr.name.startsWith('xmlns') || attr.name.includes(':')) {
            node.removeAttribute(attr.name);
          }
        }
      }
      for (let child of node.children) {
        removeNS(child);
      }
    };
    removeNS(xmlDoc.documentElement);
    return xmlDoc;
  };

  const processXML = (xmlContent, fileName) => {
    const parser = new DOMParser();
    let xmlDoc = parser.parseFromString(xmlContent, "text/xml");
    xmlDoc = removeNamespaces(xmlDoc);

    const infNFe = xmlDoc.getElementsByTagName('infNFe')[0];
    if (!infNFe) {
      console.warn(`Arquivo ${fileName} não contém tag <infNFe>`);
      return;
    }

    const ide = infNFe.getElementsByTagName('ide')[0];
    const emit = infNFe.getElementsByTagName('emit')[0];
    const dest = infNFe.getElementsByTagName('dest')[0];
    const dets = infNFe.getElementsByTagName('det');

    const numeroNF = ide?.getElementsByTagName('nNF')[0]?.textContent || '';
    const emitenteCNPJ = emit?.getElementsByTagName('CNPJ')[0]?.textContent || '';
    const emitenteNome = emit?.getElementsByTagName('xNome')[0]?.textContent || '';
    const dataEmissao = ide?.getElementsByTagName('dhEmi')[0]?.textContent || '';
    const destinatarioCNPJ = dest?.getElementsByTagName('CNPJ')[0]?.textContent || '';
    const destinatarioNome = dest?.getElementsByTagName('xNome')[0]?.textContent || '';

    Array.from(dets).forEach(produto => {
      const prod = produto.getElementsByTagName('prod')[0];
      const imposto = produto.getElementsByTagName('imposto')[0];

      const nomeProduto = prod?.getElementsByTagName('xProd')[0]?.textContent || '';
      let quantidade = prod?.getElementsByTagName('qCom')[0]?.textContent || '0';
      const valorUnitario = prod?.getElementsByTagName('vUnCom')[0]?.textContent || '0';
      const cfop = prod?.getElementsByTagName('CFOP')[0]?.textContent || '';
      const cst = imposto?.getElementsByTagName('ICMS')[0]?.children[0]?.getElementsByTagName('CST')[0]?.textContent || '';
      const ncm = prod?.getElementsByTagName('NCM')[0]?.textContent || '';

      const pisTag = imposto?.getElementsByTagName('PIS')[0];
      const cofinsTag = imposto?.getElementsByTagName('COFINS')[0];

      const pisValor = parseFloat(pisTag?.getElementsByTagName('vPIS')[0]?.textContent || '0');
      const cofinsValor = parseFloat(cofinsTag?.getElementsByTagName('vCOFINS')[0]?.textContent || '0');

      const cstPIS = pisTag?.children[0]?.getElementsByTagName('CST')[0]?.textContent || '';
      const cstCOFINS = cofinsTag?.children[0]?.getElementsByTagName('CST')[0]?.textContent || '';

      quantidade = parseFloat(quantidade.replace(',', '.')) || 0;

      const row = document.createElement('tr');
      row.innerHTML = `
        <td>${numeroNF}</td>
        <td>${emitenteCNPJ}</td>
        <td>${emitenteNome}</td>
        <td>${dataEmissao}</td>
        <td>${destinatarioCNPJ}</td>
        <td>${destinatarioNome}</td>
        <td>${nomeProduto}</td>
        <td>${ncm}</td>
        <td>${cfop}</td>
        <td>${cst}</td>
        <td>${quantidade.toFixed(2)}</td>
        <td>R$ ${parseFloat(valorUnitario).toFixed(2)}</td>
        <td>R$ ${pisValor.toFixed(2)}</td>
        <td>R$ ${cofinsValor.toFixed(2)}</td>
        <td>${cstPIS}</td>
        <td>${cstCOFINS}</td>
      `;
      tableBody.appendChild(row);
    });
  };

  const extractXMLsFromZip = async (zipObj) => {
    for (const filename of Object.keys(zipObj.files)) {
      const entry = zipObj.files[filename];
      if (entry.dir) continue;

      if (filename.endsWith('.xml')) {
        const content = await entry.async("text");
        processXML(content, filename);
      } else if (filename.endsWith('.zip')) {
        const nestedZipData = await entry.async("blob");
        const nestedZip = await JSZip.loadAsync(nestedZipData);
        await extractXMLsFromZip(nestedZip); // recursivo
      }
    }
  };

  for (const file of files) {
    if (file.name.endsWith('.zip')) {
      const zip = await JSZip.loadAsync(file);
      await extractXMLsFromZip(zip);
    } else if (file.name.endsWith('.xml')) {
      const reader = new FileReader();
      reader.onload = function(e) {
        processXML(e.target.result, file.name);
      };
      reader.readAsText(file);
    }
  }

  // Quando terminar tudo, esconder a mensagem
  document.getElementById('loadingMessage').style.display = 'none';
});

// Botão de download Excel
const downloadBtn = document.getElementById('downloadExcel');
downloadBtn.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(document.getElementById('resultTable'));
  XLSX.utils.book_append_sheet(wb, ws, 'NFe');
  XLSX.writeFile(wb, 'notas_fiscais.xlsx');
});
