function(properties, context) {

    // =========================
    //  CONSOLE VISUAL (LOGGER)
    // =========================
    function log(msg, type = "info") {
        if (!properties.console) {
            // Caso o console esteja desativado, apenas registra no console padrão
            if (type === "error") console.error(msg);
            else if (type === "warn") console.warn(msg);
            else console.log(msg);
            return;
        }

        try {
            // Cria o container se não existir
            let logger = document.getElementById("plugin-logger");
            if (!logger) {
                logger = document.createElement("div");
                logger.id = "plugin-logger";
                Object.assign(logger.style, {
                    position: "fixed",
                    bottom: "10px",
                    right: "10px",
                    width: "400px",
                    maxHeight: "300px",
                    overflowY: "auto",
                    background: "rgba(0,0,0,0.85)",
                    color: "#0f0",
                    fontFamily: "monospace",
                    fontSize: "12px",
                    padding: "10px 10px 30px 10px",
                    borderRadius: "8px",
                    zIndex: 999999,
                    boxShadow: "0 0 10px rgba(0,0,0,0.5)"
                });

                // Botão de fechar
                const closeBtn = document.createElement("button");
                closeBtn.textContent = "×";
                Object.assign(closeBtn.style, {
                    position: "absolute",
                    top: "5px",
                    right: "8px",
                    background: "rgba(255,255,255,0.1)",
                    border: "none",
                    color: "#fff",
                    fontSize: "18px",
                    cursor: "pointer",
                    borderRadius: "50%",
                    width: "22px",
                    height: "22px",
                    lineHeight: "18px",
                    textAlign: "center"
                });
                closeBtn.onclick = () => logger.remove();
                logger.appendChild(closeBtn);

                // Título
                const title = document.createElement("div");
                title.textContent = "📜 Console do Plugin";
                Object.assign(title.style, {
                    fontWeight: "bold",
                    color: "#00ffcc",
                    marginBottom: "6px"
                });
                logger.appendChild(title);

                // Área de logs
                const content = document.createElement("div");
                content.id = "plugin-log-content";
                logger.appendChild(content);

                document.body.appendChild(logger);
            }

            const content = document.getElementById("plugin-log-content");

            // Define cor por tipo
            const color = type === "error" ? "#ff5555" :
                          type === "warn" ? "#ffaa00" :
                          "#00ffcc";

            const line = document.createElement("div");
            line.style.color = color;
            line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
            content.appendChild(line);

            content.scrollTop = content.scrollHeight;
        } catch (e) {
            console.error("Erro no logger:", e);
        }
    }

    // =========================
    //  INÍCIO DO CÓDIGO
    // =========================
    log("🚀 Iniciando geração do Excel...");

    let jsonData = properties.jsonData;

    // 🔧 Decodifica JSON com limpeza preventiva
    try {
        if (typeof jsonData === 'string') {
            log("🧹 Limpando JSON...");
            jsonData = jsonData
                .replace(/[\u0000-\u001F]+/g, " ")   // remove caracteres invisíveis
                .replace(/\r?\n|\r/g, " ")           // remove quebras de linha
                .replace(/\t/g, " ");                // remove tabs

            jsonData = JSON.parse(jsonData);

            while (typeof jsonData === 'string') {
                jsonData = JSON.parse(jsonData);
            }
        }

        log("✅ JSON decodificado com sucesso.");
    } catch (error) {
        log("❌ Erro ao interpretar JSON: " + error.message, "error");
        return;
    }

    if (!Array.isArray(jsonData) || jsonData.length === 0) {
        log("❌ jsonData inválido ou vazio.", "error");
        return;
    }

    // 🧩 Verifica bibliotecas
    if (typeof XLSX === "undefined") {
        log("❌ ERRO: Biblioteca XLSX não encontrada no ambiente LIVE.", "error");
        return;
    }
    if (typeof saveAs === "undefined") {
        log("❌ ERRO: Biblioteca FileSaver (saveAs) não encontrada.", "error");
        return;
    }

    log("📚 Bibliotecas XLSX e FileSaver carregadas com sucesso.");

    // =========================
    //  CRIAÇÃO DO EXCEL
    // =========================
    const headerRow = Object.keys(jsonData[0]);
    const excelData = [headerRow, ...jsonData.map(item => headerRow.map(key => item[key]))];

    const worksheet = XLSX.utils.aoa_to_sheet(excelData);

    // Ajusta largura automaticamente
    worksheet['!cols'] = headerRow.map((_, i) => {
        const maxLength = Math.max(...excelData.map(r => (r[i] ? r[i].toString().length : 0)));
        return { wch: Math.min(maxLength + 2, 50) };
    });

    // Estilos de célula
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
            const cell = worksheet[cellRef];
            if (!cell) continue;

            const baseStyle = {
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                },
                alignment: { vertical: 'center', horizontal: R === 0 ? 'center' : 'left' }
            };

            if (R === 0) {
                baseStyle.fill = { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'D9D9D9' } };
                baseStyle.font = { bold: true, color: { rgb: '000000' } };
            }
            cell.s = baseStyle;
        }
    }

    // Cria workbook
    const workbook = XLSX.utils.book_new();
    const sheetName = properties.sheetName || properties.nomeAba || "Dados";
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    // Gera e salva arquivo
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    const fileName = (properties.fileName || "arquivo") + ".xlsx";

    try {
        saveAs(blob, fileName);
        log("✅ Arquivo Excel gerado com sucesso: " + fileName);
    } catch (err) {
        log("❌ Erro ao salvar arquivo: " + err.message, "error");
    }

    // Converter string binária em ArrayBuffer
    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    log("🏁 Processo concluído!");
}
