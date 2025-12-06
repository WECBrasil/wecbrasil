class ExtratoProcessor {
  constructor() {
    this.dataFrame = null;
    this.fileName = "";
    this.logoData = null;
    this.titleText = "Extrato Departamento Missionário Contas por Departamento";
    this.negativeRows = [];
    this.columns = [];
    this.valColName = null;
    this.dataColName = null;
    this.selectedNegativesToRemove = new Set();
    this.deptCol = "Departamento";

    this.initializeEventListeners();
  }

  initializeEventListeners() {
    document.getElementById("uploadForm").addEventListener("submit", (e) => this.handleUpload(e));
    document.getElementById("exportBtn").addEventListener("click", () => this.handleExport());
    document.getElementById("backBtn").addEventListener("click", () => this.backToUpload());
    document.getElementById("selectAllCheckbox").addEventListener("change", (e) => this.toggleSelectAll(e));
    document.getElementById("applyNegSelection").addEventListener("click", () => this.applyNegSelection());
    document.getElementById("logoInput").addEventListener("change", (e) => this.handleLogoUpload(e));

    // Neg modal trigger: usar API do Bootstrap quando disponível
    const negBtn = document.getElementById("negModalBtn");
    if (negBtn) {
      negBtn.addEventListener("click", () => {
        const modalEl = document.getElementById("negModal");
        if (!modalEl) return;
        if (window.bootstrap && typeof bootstrap.Modal === "function") {
          try {
            const inst = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl);
            inst.show();
            return;
          } catch (e) {
            console.warn("bootstrap.Modal.show falhou:", e);
          }
        }
        // fallback mínimo
        modalEl.style.display = "block";
        modalEl.classList.add("show");
        document.body.classList.add("modal-open");
      });
    }
  }

  showAlert(message, type = "danger") {
    const alertBox = document.getElementById("alertBox");
    const alert = document.createElement("div");
    alert.className = `alert alert-${type} alert-dismissible fade show`;
    alert.role = "alert";
    alert.innerHTML = `
      ${message}
      <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Fechar"></button>
    `;
    alertBox.appendChild(alert);

    setTimeout(() => {
      alert.remove();
    }, 5000);
  }

  // --- NOVOS HELPERS: normalização de valores e datas ---
  parseMoney(val) {
    if (val === null || val === undefined || val === "") return 0;
    if (typeof val === "number") return val;
    let s = String(val).trim();
    let negative = false;
    if (/^\(.*\)$/.test(s)) {
      negative = true;
      s = s.replace(/[()]/g, "");
    }
    // remover prefixos comuns e espaços
    s = s.replace(/[R$\s\xa0]/g, "");
    // se existir vírgula, trate ponto como separador de milhar e vírgula como decimal
    if (s.indexOf(",") > -1) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // caso: "1234.56" ou "1.234" -> aceitável
      s = s.replace(/,/g, ".");
    }
    const n = parseFloat(s);
    if (isNaN(n)) return 0;
    return negative ? -Math.abs(n) : n;
  }

  parseExcelDate(val) {
    // Excel serial -> JS Date (assume 1900 system)
    if (!val && val !== 0) return "";
    if (val instanceof Date) return val;
    if (typeof val === "number") {
      // Excel serial to JS date
      const ms = Math.round((val - 25569) * 86400 * 1000);
      return new Date(ms);
    }
    // se for string, tenta criar Date
    const d = new Date(String(val).trim());
    if (!isNaN(d.getTime())) return d;
    return String(val);
  }

  // --- ADICIONADAS: formatação de data e moeda (evita erros se ausentes) ---
  formatDataBrasileira(valor) {
    if (!valor && valor !== 0) return "";
    try {
      // se for número serial do Excel, converte
      if (typeof valor === "number") valor = this.parseExcelDate(valor);
      if (valor instanceof Date && !isNaN(valor.getTime())) {
        return valor.toLocaleDateString("pt-BR");
      }
      // string: tenta parsear ou retorna substring
      const s = String(valor);
      const d = new Date(s);
      if (!isNaN(d.getTime())) return d.toLocaleDateString("pt-BR");
      // se já tiver formato dd/mm/yyyy retorna até o espaço (remover hora)
      if (s.includes("/")) return s.split(" ")[0];
      return s.split(" ")[0];
    } catch {
      return String(valor);
    }
  }

  formatMoedaBrasileira(valor) {
    if (valor === null || valor === undefined || valor === "") return "";
    // garantir número
    const num = (typeof valor === "number") ? valor : this.parseMoney(valor);
    if (isNaN(num)) return "";
    return num.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  }
  // --- fim adições ---

  async handleUpload(e) {
    e.preventDefault();
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) {
      this.showAlert("Nenhum arquivo selecionado.", "danger");
      return;
    }

    this.fileName = file.name;

    try {
      const data = await file.arrayBuffer();
      // ler com cellDates para preservar datas e tipo bruto
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

      if (jsonData.length < 2) {
        this.showAlert("Arquivo inválido. Mínimo 2 linhas necessárias.", "danger");
        return;
      }

      this.titleText = String(jsonData[0][0] || this.titleText).trim();
      this.columns = jsonData[1].map((c) => String(c || "").trim()).filter((c) => c !== "");
      this.dataFrame = jsonData.slice(2).map((row) => {
        const obj = {};
        this.columns.forEach((col, idx) => {
          obj[col] = row[idx] !== undefined ? row[idx] : "";
        });
        return obj;
      });

      // remover linhas vazias
      this.dataFrame = this.dataFrame.filter((row) =>
        Object.values(row).some((val) => val !== "" && val !== null && val !== undefined)
      );

      // detectar colunas (mais abrangente)
      this.valColName = this.findColumn(["valor da conta", "valor", "vlr", "vlor", "valor_material", "valor_pago"]);
      this.dataColName = this.findColumn(["data de pagto", "data de recbto", "data", "dat", "dt"]);
      // converter coluna de valores para número (robusto)
      if (this.valColName) {
        this.dataFrame.forEach((row) => {
          row[this.valColName] = this.parseMoney(row[this.valColName]);
        });
      }

      // converter colunas de data (se necessário)
      if (this.dataColName) {
        this.dataFrame.forEach((row) => {
          const v = row[this.dataColName];
          const parsed = this.parseExcelDate(v);
          row[this.dataColName] = parsed instanceof Date ? parsed : parsed;
        });
      }

      this.findNegativeRows();
      this.showPreview();
      this.showAlert("Arquivo importado com sucesso!", "success");
    } catch (error) {
      this.showAlert(`Erro ao ler arquivo: ${error.message}`, "danger");
      console.error(error);
    }
  }

  findColumn(keywords) {
    for (const col of this.columns) {
      const colLower = col.toLowerCase();
      for (const keyword of keywords) {
        if (colLower.includes(keyword)) {
          return col;
        }
      }
    }
    return null;
  }

  findNegativeRows() {
    this.negativeRows = [];
    if (!this.valColName) return;

    this.dataFrame.forEach((row, idx) => {
      const val = row[this.valColName];
      if (typeof val === "number" && val < 0) {
        const deptVal = row["Departamento"] || row[this.findColumn(["departamento"])] || "";
        this.negativeRows.push({
          index: idx,
          data: this.dataColName ? this.formatDataBrasileira(row[this.dataColName]) : "",
          dept: String(deptVal).trim(),
          value: this.formatMoedaBrasileira(val),
          rowValue: val,
        });
      }
    });
  }

  showPreview() {
    document.getElementById("uploadSection").classList.add("hidden");
    document.getElementById("previewSection").classList.remove("hidden");
    document.getElementById("previewFileName").textContent = this.fileName;
    document.getElementById("valColDetected").textContent = this.valColName || "Nenhuma";
    document.getElementById("negCountDisplay").textContent = this.negativeRows.length;
    document.getElementById("negCountModal").textContent = this.negativeRows.length;

    this.selectedNegativesToRemove.clear();
    document.getElementById("remove_idxs").value = "";

    this.renderPreviewTable();
    this.renderNegativeModal();
    window.scrollTo(0, 0);
  }

  renderPreviewTable() {
    const headerRow = document.getElementById("previewHeader");
    const tbody = document.getElementById("previewBody");

    headerRow.innerHTML = this.columns.map((col) => `<th>${col}</th>`).join("");

    tbody.innerHTML = this.dataFrame
      .map((row, idx) => {
        const isNegative = this.negativeRows.some((nr) => nr.index === idx);
        const rowClass = isNegative ? "negative-row" : "";
        const cells = this.columns
          .map((col) => {
            const val = row[col];
            const cellClass = isNegative && col === this.valColName ? "negative-cell" : "";
            let displayVal = val;

            if (col === this.dataColName && val) {
              displayVal = this.formatDataBrasileira(val);
            } else if (col === this.valColName && typeof val === "number") {
              displayVal = this.formatMoedaBrasileira(val);
            }

            return `<td class="${cellClass}">${displayVal}</td>`;
          })
          .join("");

        return `<tr class="${rowClass}">${cells}</tr>`;
      })
      .join("");
  }

  renderNegativeModal() {
    const tbody = document.getElementById("negTableBody");
    const negTable = document.getElementById("negTable");
    const emptyMsg = document.getElementById("modalEmptyMsg");
    const selectAllCheckbox = document.getElementById("selectAllCheckbox");

    if (this.negativeRows.length === 0) {
      negTable.style.display = "none";
      emptyMsg.style.display = "block";
      selectAllCheckbox.checked = false;
      return;
    }

    negTable.style.display = "table";
    emptyMsg.style.display = "none";

    tbody.innerHTML = this.negativeRows
      .map(
        (nr) => `
      <tr class="table-danger">
        <td>
          <input type="checkbox" class="neg-check form-check-input" data-idx="${nr.index}">
        </td>
        <td><strong>${nr.index}</strong></td>
        <td>${nr.data}</td>
        <td>${nr.dept}</td>
        <td><span class="text-danger font-weight-bold">${nr.value}</span></td>
      </tr>
    `
      )
      .join("");

    selectAllCheckbox.checked = false;
  }

  toggleSelectAll(e) {
    const checks = document.querySelectorAll(".neg-check");
    checks.forEach((cb) => (cb.checked = e.target.checked));
  }

  // === Helper robusto para fechar modal e limpar backdrop ===
  closeModal(modalEl) {
    try {
      if (!modalEl) return;
      // Preferir API bootstrap
      if (window.bootstrap && typeof bootstrap.Modal === "function") {
        try {
          const inst = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl);
          inst.hide();
        } catch (e) {
          console.warn("closeModal: erro ao usar bootstrap.Modal.hide:", e);
          modalEl.classList.remove("show");
          modalEl.style.display = "none";
        }
      } else {
        modalEl.classList.remove("show");
        modalEl.style.display = "none";
        modalEl.setAttribute("aria-hidden", "true");
        modalEl.removeAttribute("aria-modal");
      }
    } catch (e) {
      console.warn("closeModal: fallback de fechamento falhou:", e);
      try { modalEl.classList.remove("show"); modalEl.style.display = "none"; } catch {}
    } finally {
      // Limpeza SEGURA de backdrops: remover apenas aqueles com z-index elevado (backdrops do bootstrap)
      try {
        document.querySelectorAll(".modal-backdrop").forEach((el) => {
          const zi = parseInt(window.getComputedStyle(el).zIndex || "0", 10);
          if (!isNaN(zi) && zi >= 1000) el.remove();
        });
      } catch (cleanupErr) {
        // última tentativa: remover todos backdrops (menos preferível)
        document.querySelectorAll(".modal-backdrop").forEach((el) => el.remove());
      }

      // restaurar estado do body (não remove estilos além do necessário)
      document.body.classList.remove("modal-open");
      // remover overflow/scroll lock e padding-right adicionados pelo bootstrap
      try { document.body.style.removeProperty("overflow"); } catch {}
      try { document.body.style.removeProperty("padding-right"); } catch {}

      // garantir foco em botão principal (ajuda UX)
      const negBtn = document.getElementById("negModalBtn");
      if (negBtn) try { negBtn.focus(); } catch {}
    }
  }

  applyNegSelection() {
    try {
      const checks = document.querySelectorAll(".neg-check:checked");
      const idxs = Array.from(checks)
        .map((cb) => parseInt(cb.getAttribute("data-idx"), 10))
        .filter((n) => !isNaN(n));

      this.selectedNegativesToRemove = new Set(idxs);
      document.getElementById("remove_idxs").value = Array.from(idxs).join(",");

      const modalEl = document.getElementById("negModal");

      // fechar modal imediatamente (não aguardar remoção pesada)
      this.closeModal(modalEl);

      // Fazer a remoção de forma assíncrona pequena para liberar o thread UI
      setTimeout(() => {
        try {
          if (this.selectedNegativesToRemove.size === 0) {
            this.showAlert("Nenhum registro selecionado para remoção.", "info");
            return;
          }

          // filtrar dataFrame por índices (índices referenciam this.dataFrame atual)
          const toRemove = this.selectedNegativesToRemove;
          this.dataFrame = this.dataFrame.filter((_, idx) => !toRemove.has(idx));

          // recomputar negativos e atualizar UI
          this.findNegativeRows();
          this.renderPreviewTable();
          this.renderNegativeModal();
          this.selectedNegativesToRemove.clear();
          document.getElementById("remove_idxs").value = "";

          this.showAlert(`${toRemove.size} registro(s) removido(s).`, "warning");
        } catch (innerErr) {
          console.error("Erro ao aplicar remoção assíncrona:", innerErr);
          this.showAlert("Erro ao remover registros. Veja console.", "danger");
        }
      }, 50);
    } catch (err) {
      console.error("applyNegSelection erro:", err);
      // tentativa de fechar modal e desbloquear UI mesmo em erro
      try { this.closeModal(document.getElementById("negModal")); } catch {}
      this.showAlert("Erro inesperado ao processar seleção.", "danger");
    }
  }

  async handleExport() {
    const deptCol = document.getElementById("deptColInput").value.trim();

    if (!deptCol) {
      this.showAlert("Nome da coluna de referência é obrigatório.", "danger");
      return;
    }

    if (!this.columns.includes(deptCol)) {
      this.showAlert(`Coluna "${deptCol}" não encontrada na planilha.`, "danger");
      return;
    }

    this.deptCol = deptCol;

    let filteredData = this.dataFrame;
    if (this.selectedNegativesToRemove.size > 0) {
      filteredData = this.dataFrame.filter(
        (_, idx) => !this.selectedNegativesToRemove.has(idx)
      );
    }

    try {
      this.showAlert("🔄 Gerando PDFs... Por favor aguarde.", "info");
      const zipContent = await this.generatePDFZip(filteredData, deptCol);
      this.downloadZip(zipContent);
      this.showAlert("✅ PDFs gerados com sucesso!", "success");
    } catch (error) {
      this.showAlert(`❌ Erro ao gerar PDFs: ${error.message}`, "danger");
      console.error(error);
    }
  }

  async generatePDFZip(data, deptCol) {
    if (typeof JSZip === "undefined") {
      throw new Error("JSZip não encontrado (vendor/jszip.min.js).");
    }
    const zip = new JSZip();

    // Agrupar por departamento
    const grouped = {};
    data.forEach((row) => {
      const dept = String(row[deptCol] || "Sem Departamento").trim();
      if (!grouped[dept]) grouped[dept] = [];
      grouped[dept].push(row);
    });

    console.log(`📑 Gerando ${Object.keys(grouped).length} PDF(s)...`);

    const depts = Object.keys(grouped).sort();
    const failed = [];

    for (const dept of depts) {
      const deptData = grouped[dept];
      try {
        const pdfBlob = await this.generatePDFContent(deptData, dept);
        if (!pdfBlob || !(pdfBlob instanceof Blob) || pdfBlob.size === 0) {
          throw new Error("PDF gerado inválido/vazio");
        }
        const fileName = this.sanitizeFileName(`${dept} - Extrato.pdf`);
        zip.file(fileName, pdfBlob);
        console.log(`✅ PDF adicionado ao ZIP: ${fileName} (${pdfBlob.size} bytes)`);
      } catch (err) {
        console.error(`❌ Falha ao gerar PDF para "${dept}":`, err);
        // fallback: gerar PDF simples com texto para não perder o departamento
        try {
          console.log(`🔁 Tentando fallback textual para "${dept}"`);
          const fallbackBlob = await (async () => {
            // tenta criar um PDF mínimo com jsPDF (se disponível)
            const jsPDFConstructor = (typeof jsPDF !== "undefined") ? jsPDF : (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : null;
            if (!jsPDFConstructor) {
              // criar blob de texto simples se jsPDF não disponível
              const txt = `Departamento: ${dept}\n\n${JSON.stringify(deptData, null, 2)}`;
              return new Blob([txt], { type: "text/plain" });
            }
            const pdf = new jsPDFConstructor({ orientation: "landscape", unit: "mm", format: "a4" });
            const lines = JSON.stringify(deptData, null, 2).split("\n");
            let y = 10;
            pdf.setFontSize(10);
            pdf.text(`Extrato - ${dept}`, 10, y);
            y += 8;
            for (const line of lines) {
              pdf.text(line.slice(0, 120), 10, y);
              y += 6;
              if (y > 280) { pdf.addPage(); y = 10; }
            }
            return pdf.output("blob");
          })();
          const fileName = this.sanitizeFileName(`${dept} - Extrato (fallback).pdf`);
          zip.file(fileName, fallbackBlob);
          console.log(`✅ Fallback adicionado: ${fileName} (${fallbackBlob.size} bytes)`);
        } catch (fbErr) {
          console.error(`❌ Fallback falhou para "${dept}":`, fbErr);
          failed.push(dept);
        }
      }
    }

    if (Object.keys(zip.files).length === 0) {
      throw new Error("Nenhum PDF válido gerado. ZIP vazio.");
    }

    if (failed.length) {
      console.warn("Alguns departamentos falharam e foram omitidos:", failed);
    }

    return await zip.generateAsync({ type: "blob" });
  }

  async generatePDFContent(deptData, deptName) {
    try {
      const jsPDFCtor =
        (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF :
        (typeof jsPDF !== "undefined" ? jsPDF : null);

      if (!jsPDFCtor) return this._generateTextualPdfBlob(deptData, deptName);

      // landscape A4
      const doc = new jsPDFCtor({ orientation: "landscape", unit: "mm", format: "a4" });
      const pageWidth = doc.internal.pageSize.getWidth();
      const margin = 10;
      const headerH = 28;

      // Header background — cor solicitada: #140851 => rgb(20,8,81)
      doc.setFillColor(20, 8, 81);
      doc.rect(0, 0, pageWidth, headerH, "F");

      // Logo (se existente) — tentar PNG/JPEG
      if (this.logoData) {
        try {
          const logoW = 20;
          const logoH = headerH - 8;
          doc.addImage(this.logoData, (this.logoData.indexOf("png") > -1 ? "PNG" : "JPEG"), margin, 4, logoW, logoH);
        } catch (e) {
          console.warn("Logo não adicionada:", e);
        }
      }

      // Título e subtítulo no header (branco)
      doc.setTextColor(255, 255, 255);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14);
      const titleX = margin + (this.logoData ? 40 : 0);
      doc.text(this.titleText || "Extrato", titleX, 9);

      doc.setFontSize(11);
      doc.setFont("helvetica", "normal");
      doc.text(`Departamento: ${deptName}`, titleX, 16);

      // Preparar dados para autoTable
      const head = [ this.columns.map(c => String(c)) ];

      // calcular total do extrato (coluna de valores)
      const valIdx = this.columns.indexOf(this.valColName);
      let totalValor = 0;
      const body = deptData.map(row => {
        const line = this.columns.map(col => {
          if (col === this.dataColName) return this.formatDataBrasileira(row[col]);
          if (col === this.valColName) {
            const num = (typeof row[col] === "number") ? row[col] : this.parseMoney(row[col]);
            if (!isNaN(num)) totalValor += Number(num);
            return this.formatMoedaBrasileira(num);
          }
          return (row[col] === null || row[col] === undefined) ? "" : String(row[col]);
        });
        return line;
      });

      // incluir linha de TOTAL ao final (colocar "TOTAL" na primeira coluna e valor na coluna de valores)
      const totalRow = new Array(this.columns.length).fill("");
      if (this.columns.length > 0) totalRow[0] = "TOTAL";
      if (valIdx >= 0) totalRow[valIdx] = this.formatMoedaBrasileira(totalValor);
      body.push(totalRow);

      // Forçar estilos legíveis (preto / cinza) e layout amplo
      const startY = headerH + 6;
      if (typeof doc.autoTable === "function") {
        doc.autoTable({
          head,
          body,
          startY,
          margin: { left: margin, right: margin },
          styles: {
            font: "helvetica",
            fontSize: 9,
            textColor: [30, 30, 30],
            cellPadding: 4,
            overflow: "linebreak"
          },
          headStyles: {
            fillColor: [245, 245, 245],
            textColor: [20, 20, 20],
            fontStyle: "bold"
          },
          alternateRowStyles: {
            fillColor: [250, 250, 250]
          },
          tableLineWidth: 0.12,
          tableLineColor: [200, 200, 200],
          theme: "striped",
          didParseCell: (data) => {
            // estilizar a última linha (TOTAL) para destaque
            if (data.row.section === 'body' && data.row.index === body.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.textColor = [10, 10, 10];
              data.cell.styles.fillColor = [235, 235, 250]; // leve tom para total
            }
          }
        });

        // rodapé com data de geração
        const finalY = doc.lastAutoTable ? doc.lastAutoTable.finalY + 8 : startY + 6;
        doc.setFontSize(9);
        doc.setTextColor(100);
        doc.text(`Gerado em ${new Date().toLocaleString('pt-BR')}`, pageWidth - margin, finalY, { align: "right" });

        const blob = doc.output("blob");
        if (!blob || blob.size < 2000) {
          console.warn("PDF gerado pequeno/suspeito, fallback textual:", blob ? blob.size : 0);
          return this._generateTextualPdfBlob(deptData, deptName);
        }
        return blob;
      }

      // AutoTable não disponível -> fallback textual via jsPDF
      doc.setFontSize(11);
      doc.setTextColor(0);
      let y = startY;
      const colCount = this.columns.length || 1;
      const usableWidth = pageWidth - margin * 2;
      const colWidth = Math.max(30, Math.floor(usableWidth / Math.min(colCount, 6)));

      // header line
      doc.setFont("helvetica", "bold");
      this.columns.forEach((c, i) => {
        doc.text(String(c).slice(0, 20), margin + i * colWidth, y);
      });
      y += 6;
      doc.setFont("helvetica", "normal");
      for (let rIndex = 0; rIndex < body.length; rIndex++) {
        const r = body[rIndex];
        let x = margin;
        for (let i = 0; i < r.length; i++) {
          doc.text(String(r[i]).slice(0, 40), x, y);
          x += colWidth;
        }
        y += 6;
        if (y > doc.internal.pageSize.getHeight() - 20) {
          doc.addPage();
          y = margin;
        }
      }

      return doc.output("blob");
    } catch (err) {
      console.error("Erro em generatePDFContent:", err);
      return this._generateTextualPdfBlob(deptData, deptName);
    }
  }

  // helper reutilizável de fallback textual (usa jsPDF se houver)
  _generateTextualPdfBlob(deptDataInner, deptNameInner) {
    const jsPDFConstructor = (typeof jsPDF !== "undefined") ? jsPDF : (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : null;
    if (!jsPDFConstructor) {
      const txt = `Extrato - ${deptNameInner}\n\n` + JSON.stringify(deptDataInner, null, 2);
      return new Blob([txt], { type: "text/plain" });
    }
    const pdf = new jsPDFConstructor({ orientation: "portrait", unit: "mm", format: "a4" });
    const left = 10;
    let y = 12;
    pdf.setFontSize(12);
    pdf.text(`${this.titleText}`, left, y);
    y += 7;
    pdf.setFontSize(11);
    pdf.text(`Departamento: ${deptNameInner}`, left, y);
    y += 8;
    pdf.setFontSize(9);

    const colCount = this.columns.length || 1;
    const pageWidth = pdf.internal.pageSize.getWidth();
    const usableWidth = pageWidth - left * 2;
    const colWidth = Math.max(30, Math.floor(usableWidth / Math.min(colCount, 5)));

    let x = left;
    this.columns.forEach((col) => {
      pdf.text(String(col).slice(0, 15), x, y);
      x += colWidth;
    });
    y += 6;

    for (const row of deptDataInner) {
      x = left;
      for (const col of this.columns) {
        let val = row[col];
        if (col === this.dataColName && val) val = this.formatDataBrasileira(val);
        else if (col === this.valColName) {
          const num = (typeof val === "number") ? val : this.parseMoney(val);
          val = this.formatMoedaBrasileira(num);
        } else if (val === null || val === undefined) val = "";
        pdf.text(String(val).slice(0, 40), x, y);
        x += colWidth;
      }
      y += 6;
      if (y > pdf.internal.pageSize.getHeight() - 20) {
        pdf.addPage();
        y = 12;
      }
    }

    return pdf.output("blob");
  }

  // === ADICIONADO: gera HTML limpo (sem <html>/<body>) para cada departamento ===
  generateHTMLForPDF(deptData, deptName) {
    let html = `<div style="font-family: Arial, sans-serif; font-size: 11px; color:#000;">`;
    // header
    html += `<div style="text-align:center;margin-bottom:12px;">`;
    if (this.logoData) {
      html += `<div><img src="${this.logoData}" style="max-width:60px;max-height:60px;margin-bottom:8px;" /></div>`;
    }
    html += `<div style="font-weight:700;font-size:14px;margin-bottom:2px;">${this.titleText}</div>`;
    html += `<div style="font-size:12px;color:#333;margin-bottom:4px;">${deptName}</div>`;
    html += `</div>`;

    // table header
    html += `<table style="width:100%;border-collapse:collapse;"><thead><tr>`;
    this.columns.forEach((col) => {
      html += `<th style="text-align:left;padding:6px;border:1px solid #ddd;background:#f5f5f5;font-size:10px;">${col}</th>`;
    });
    html += `</tr></thead><tbody>`;

    // rows
    let totalValor = 0;
    deptData.forEach((row) => {
      html += `<tr>`;
      this.columns.forEach((col) => {
        let val = row[col];
        let display = "";
        if (col === this.dataColName && val) {
          display = this.formatDataBrasileira(val);
        } else if (col === this.valColName) {
          const num = (typeof val === "number") ? val : this.parseMoney(val);
          if (typeof num === "number" && !isNaN(num)) totalValor += num;
          display = this.formatMoedaBrasileira(num);
        } else {
          display = (val === null || val === undefined) ? "" : String(val);
        }
        html += `<td style="padding:6px;border:1px solid #ddd;font-size:10px;">${display}</td>`;
      });
      html += `</tr>`;
    });

    // total row
    html += `<tr style="font-weight:700;background:#eef;">`;
    this.columns.forEach((col, idx) => {
      if (idx === 0) {
        html += `<td style="padding:6px;border:1px solid #999;">TOTAL</td>`;
      } else if (col === this.valColName) {
        html += `<td style="padding:6px;border:1px solid #999;">${this.formatMoedaBrasileira(totalValor)}</td>`;
      } else {
        html += `<td style="padding:6px;border:1px solid #999;"></td>`;
      }
    });
    html += `</tr>`;

    html += `</tbody></table>`;

    html += `<div style="text-align:center;margin-top:10px;font-size:9px;color:#666;">Gerado em ${new Date().toLocaleString('pt-BR')}</div>`;
    html += `</div>`;

    return html;
  }

  // === ADICIONADO: sanitiza nome de arquivo ===
  sanitizeFileName(name) {
    return String(name).replace(/[\/\\:*?"<>|]/g, "_").trim();
  }

  downloadZip(blob) {
    if (!blob || !(blob instanceof Blob)) {
      this.showAlert("Erro: conteúdo do ZIP inválido.", "danger");
      console.error("downloadZip recebido blob inválido:", blob);
      return;
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `extratos_export_${new Date().getTime()}.zip`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  async handleLogoUpload(e) {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        this.logoData = event.target.result;
        this.showAlert("Logo carregada com sucesso!", "success");
      };
      reader.onerror = () => {
        this.showAlert("Erro ao carregar logo", "danger");
      };
      reader.readAsDataURL(file);
    }
  }

  // Carrega logo padrão automaticamente se não houver upload
  async loadDefaultLogo() {
    if (this.logoData) return; // já tem logo do upload
    const candidates = [
      '../images/wec-logo-extrato.png',
      'images/wec-logo-extrato.png',
      './images/wec-logo-extrato.png'
    ];
    for (const url of candidates) {
      try {
        const resp = await fetch(url, { cache: 'no-store' });
        if (!resp.ok) continue;
        const blob = await resp.blob();
        const reader = new FileReader();
        reader.onload = () => {
          if (reader.result) this.logoData = reader.result;
        };
        reader.readAsDataURL(blob);
        return;
      } catch (e) {
        continue;
      }
    }
  }

  backToUpload() {
    document.getElementById("previewSection").classList.add("hidden");
    document.getElementById("uploadSection").classList.remove("hidden");
    document.getElementById("fileInput").value = "";
    document.getElementById("logoInput").value = "";
    document.getElementById("deptColInput").value = "Departamento";
    this.dataFrame = null;
    this.negativeRows = [];
    this.selectedNegativesToRemove.clear();
    this.logoData = null;
    window.scrollTo(0, 0);
  }

  // Aguarda variável global disponível (html2pdf/html2canvas/jsPDF etc)
  waitForGlobal(name, timeout = 3000) {
    return new Promise((resolve, reject) => {
      if (typeof window === "undefined") return reject(new Error("window não definido"));
      if (window[name]) return resolve(window[name]);
      const start = Date.now();
      const iv = setInterval(() => {
        if (window[name]) {
          clearInterval(iv);
          return resolve(window[name]);
        }
        if (Date.now() - start > timeout) {
          clearInterval(iv);
          return reject(new Error(`${name} não encontrado dentro de ${timeout}ms`));
        }
      }, 50);
    });
  }
}

// Inicializar aplicação quando o DOM está pronto
document.addEventListener("DOMContentLoaded", () => {
  if (typeof XLSX === "undefined") {
    console.error("Erro crítico: XLSX não carregado. Verifique vendor/xlsx.full.min.js e index.html.");
    document.body.innerHTML = `<div class="m-4 alert alert-danger">Erro: biblioteca <strong>XLSX</strong> não carregada. Coloque <code>vendor/xlsx.full.min.js</code> e atualize a página.</div>`;
    return;
  }
  const processor = new ExtratoProcessor();
  // Carregar logo padrão assim que a app inicia
  processor.loadDefaultLogo().catch(() => {});
});