// Dicionário de operadores
const variavel_operadores = {
      "MANHÃ": [
        "ACLECIO FERNADO MELO", "ALECIA CRISTINA EVYNA SANTOS DE JESUS", "ALISSON BRENO CUNHA DE BARROS",
        "ALLANYA GABRIELLA DE ALMEIDA SOUSA", "ANA ACACIA ARAGAO BONFIM", "BEATRIZ VITORIA NASCIMENTO DA SILVA",
        "DIOGO SANTOS MACHADO", "DEBORA MESSIAS SANTOS", "ELAINE PINTO DE CARVALHO FREIRE",
        "JONAS LUCAS DOS SANTOS", "KELLY CRISTINA TELES DOS SANTOS",
        "LUCAS MICHAEL SANTA RITA DOS SANTOS", "LUIZ MIGUEL PATRICIO FELIX", "MARCLYS ANGELICA FERREIRA SANTOS",
        "MARICLEIDE DOS SANTOS DE SOUZA", "MONIQUE ALVES DA SILVA SANTOS",
        "ROSE KATIUSKA DOS SANTOS BIGI", "ROZEANE OLIVEIRA DOS SANTOS",
        "NATALIA FARIAS SANTOS", "ADRIANA BATISTA SILVA", "MICHELE DOS SANTOS TAVARES",
        "ISRAEL SANTOS DA SILVA", "LAIS DANIELLE HONORATO ALVES SENA",
        "TAYS MILENA BISPO DOS SANTOS", "MARIA FLORENTINA MELO SANTOS",
        "LUSINEIDE SANTIAGO SANTOS", "LARISSA VITOR DA SILVA", "BEATRIZ DE ANDRADE SANTANA",
        "KEMELLY KAROLAINNY DOS SANTOS SILVA", "LUCIANE DOS SANTOS SILVA",
        "NAIARA ANDRADE SANTOS", "SILMARA SANTOS DE JESUS",
        "CRISLAINE BORGES DOS SANTOS NASCIMENTO", "HERSLANDER JORGE DOS SANTOS",
        "LORENA FERREIRA SANTOS", "ALESSANDRA SANTANA SANTOS", "JANISSON SANTANA SANTOS"
    ],
    "TARDE": [
        "ANA RAQUEL RIBEIRO DE OLIVEIRA SANTOS", "DAYANE NUNES VICENTE", "JULIA CONCEICAO SILVA SANTOS",
        "KETLYN JULIANE SILVA BATISTA", "TALITA PRISCILA FARIAS SANTOS", "DANIELLE FERNANDES BARROS SANTOS",
        "VANESSA DE OLIVEIRA GONZAGA",
        "ACUCENA DA SILVA MARCOLINO ANDRADE", "ALISSON DA CRUZ FERREIRA",
        "GABRIELLE SALOMAO DOS SANTOS", "ROBERTH SANTOS CONCEICAO",
        "TACIANE GRACIELE INACIO DOS SANTOS", "LARISSA LUIZA VIEIRA MOTA", "ACACIA DOS SANTOS CONCEICAO", 
        "ANNA BEATRIZ OLIVEIRA FONTES",
        "JAILA RODRIGUES DOS SANTOS", "EVELYN DE FREITAS SANTOS",
        "JENNIFER CAMPOS BENIGNO", "JULIANE BARROS DA SILVA",
        "LORENA MARTINS DE SANTANA", "ANA VIRGINIA SANTOS DANTAS",
        "BRUNA CAROLINA DOS SANTOS MATOS", "PAULA ROBERTA DE JESUS SANTOS",
        "VYVYAN ADRYENE SANTOS DA SILVA"
    ],
    "INTERMEDIÁRIO": [
        "GLILMA MARIA FARIAS DE MENEZES ISMERIM", "LARISSA XAVIER DE MELO",
        "RAELLY SANTOS CALDAS LIMA", "RAYARA SANTANA LIMA", "RANYSSA RAYARA DOS SANTOS", "YASMIN SANTOS DA CONCEICAO"
    ],
    "RECEPTIVO": [
        "LEONICE PAIXAO BATISTA", "EDSON XAVIER DE BRITO", "EDUARDO MOURA SANTOS", "RAINELDES GUILHERME DOS SANTOS"
    ]
    
    };


// Normaliza nomes
function normalizeName(name) {
  return name ? name.toString().trim().toUpperCase() : "";
}

// Processa a planilha
async function processar() {
  const input = document.getElementById("planilha");
  if (!input.files.length) return alert("Selecione um arquivo!");

  const file = input.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet);

  const operadores_por_turno = variavel_operadores;

  const resultados = {};
  let totalGeral = 0;

  for (let turno in operadores_por_turno) {
    resultados[turno] = {};
    const operadores = operadores_por_turno[turno].map(normalizeName);

    operadores.forEach(op => {
      const count = json.filter(r => normalizeName(r["NOME SOLICITANTE"]) === op).length;
      if (count > 0) {
        resultados[turno][op] = count;
        totalGeral += count;
      }
    });
  }

  mostrarResultados(resultados, totalGeral);
  adicionarBotaoDownload(resultados);
}

// Renderiza os resultados em cards coloridos
function mostrarResultados(resultados, totalGeral) {
  const div = document.getElementById("resultado");
  div.innerHTML = `<div class="alert alert-success">✅ Total Geral de Acordos: ${totalGeral}</div>`;

  for (let turno in resultados) {
    const operadores = resultados[turno];

    const card = document.createElement("div");
    card.className = `turno-card turno-${turno}`;

    const header = document.createElement("div");
    header.className = `turno-header turno-${turno}`;
    header.textContent = turno;
    card.appendChild(header);

    const ul = document.createElement("ul");
    ul.className = "list-group list-group-flush";

    let totalTurno = 0;
    for (let operador in operadores) {
      const qtd = operadores[operador];
      totalTurno += qtd;

      const li = document.createElement("li");
      li.className = "list-group-item";
      li.textContent = operador;

      const badge = document.createElement("span");
      badge.className = "badge bg-primary rounded-pill";
      badge.textContent = qtd;
      li.appendChild(badge);

      ul.appendChild(li);
    }

    card.appendChild(ul);

    const totalP = document.createElement("p");
    totalP.className = "fw-bold mt-2 text-end me-2";
    totalP.textContent = `Total do turno: ${totalTurno}`;
    card.appendChild(totalP);

    div.appendChild(card);
  }
}

// Gera Excel
function gerarExcel(resultados) {
  const wb = XLSX.utils.book_new();

  for (let turno in resultados) {
    const operadores = resultados[turno];
    const data = [["Operador", "Acordos"]];

    for (let op in operadores) {
      data.push([op, operadores[op]]);
    }

    const totalTurno = Object.values(operadores).reduce((a,b)=>a+b,0);
    data.push(["Total do turno", totalTurno]);

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, turno);
  }

  XLSX.writeFile(wb, "acordos.xlsx");
}

// Botão de download
function adicionarBotaoDownload(resultados) {
  const div = document.getElementById("resultado");

  const btnAntigo = document.getElementById("btnDownload");
  if (btnAntigo) btnAntigo.remove();

  const btn = document.createElement("button");
  btn.id = "btnDownload";
  btn.className = "btn btn-success mt-3";
  btn.innerHTML = "📥 Baixar Excel";
  btn.onclick = () => gerarExcel(resultados);

  div.appendChild(btn);
}

// Evento do botão analisar
document.getElementById("form").addEventListener("submit", async (e)=>{
  e.preventDefault();
  const btn = document.getElementById("btnAnalisar");
  btn.disabled = true;
  btn.innerHTML = `<span class="loading">⏳ Processando...</span>`;
  await processar();
  btn.disabled = false;
  btn.innerHTML = `<i class="bi bi-search"></i> Analisar`;
});
