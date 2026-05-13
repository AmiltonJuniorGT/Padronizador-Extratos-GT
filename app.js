/* Código Normatização */

const $ = (id) => document.getElementById(id);
        <td>${r.data.toLocaleDateString()}</td>
        <td>${r.descricao}</td>
        <td>${r.valor.toFixed(2)}</td>
      </tr>
    `;
  }

  html += `</tbody></table>`;

  $("rawPreview").innerHTML = html;
}

function renderPivot(){

  const rows = PIVOT_DATA.slice(0,20);

  if(!rows.length){
    return;
  }

  const cols = Object.keys(rows[0]);

  let html = `
    <table>
      <thead>
        <tr>
  `;

  for(const c of cols){
    html += `<th>${c}</th>`;
  }

  html += `</tr></thead><tbody>`;

  for(const r of rows){

    html += `<tr>`;

    for(const c of cols){
      html += `<td>${r[c]}</td>`;
    }

    html += `</tr>`;
  }

  html += `</tbody></table>`;

  $("pivotPreview").innerHTML = html;
}

function exportWorkbook(){

  if(!RAW_DATA.length){
    alert("Nada processado");
    return;
  }

  const wb = XLSX.utils.book_new();

  const rawSheet = XLSX.utils.json_to_sheet(RAW_DATA);
  const pivotSheet = XLSX.utils.json_to_sheet(PIVOT_DATA);

  XLSX.utils.book_append_sheet(wb, rawSheet, "RAW_EXTRATO");
  XLSX.utils.book_append_sheet(wb, pivotSheet, "PIVOT_MENSAL");

  XLSX.writeFile(wb, "extrato_padronizado.xlsx");
}

function setStatus(msg){
  $("status").innerText = msg;
}
