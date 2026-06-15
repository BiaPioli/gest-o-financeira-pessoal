function atualizarBaseDados() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaOrigem = ss.getSheetByName("Despesas");
  const abaDestino = ss.getSheetByName("Planilha dash");
  const abaControle = ss.getSheetByName("Controle");

  abaDestino.clearContents();

  abaDestino.appendRow(["data", "mês", "ano", "tipo", "categoria", "valor"]);

  const meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
  ];

  const receitas = [
    ["Mesada", 5],
    ["Renda extra", 6],
    ["Investimento viagem", 7],
    ["Saldo anterior", 8],
    ["Salário", 9],
    ["Resgate investimento", 10]
  ];

  const investimentos = [
    ["Viagem", 16],
    ["Reserva", 17],
    ["Presente", 18]
  ];

  const despesas = [
    ["Crédito", 24],
    ["Alimentação", 25],
    ["Mercado", 26],
    ["Compras", 27],
    ["Festa/date", 28],
    ["Transporte", 29],
    ["Farmácia", 30],
    ["Extras", 31]
  ];

  for (let i = 0; i < meses.length; i++) {

    const coluna = i + 2;
    const data = new Date(2026, i, 1);

    receitas.forEach(item => {

      const valor = abaOrigem.getRange(item[1], coluna).getValue();

      abaDestino.appendRow([
        data,
        meses[i],
        2026,
        "receita",
        item[0],
        valor
      ]);
    });

    investimentos.forEach(item => {

      const valor = abaOrigem.getRange(item[1], coluna).getValue();

      abaDestino.appendRow([
        data,
        meses[i],
        2026,
        "investimento",
        item[0],
        valor
      ]);
    });

    despesas.forEach(item => {

      const valor = abaOrigem.getRange(item[1], coluna).getValue();

      abaDestino.appendRow([
        data,
        meses[i],
        2026,
        "despesa",
        item[0],
        valor
      ]);
    });
  }

  // Atualiza data e hora da última execução
abaControle.clearContents();
abaControle.getRange("A1").setValue("ultima_atualizacao");
abaControle.getRange("A2").setValue(new Date());
