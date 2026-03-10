process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

import { stripMentionsText } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import * as xlsx from "xlsx";
import * as path from "path";

const app = new App();

const normalizar = (texto: string): string => {
  return texto ? texto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "").replace(/\s+/g, " ").trim() : "";
};

app.on("message", async (context) => {
  const rawText = stripMentionsText(context.activity) || "";
  const searchText = normalizar(rawText);

  console.log(`\n--- Búsqueda recibida: "${rawText}" ---`);

  try {
    const filePath = path.join(process.cwd(), "ObrasPlantaTabla.xlsx");
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows: any[][] = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`Filas en Excel: ${rows.length}`);

    let encontrado = false;
    let respuestaTexto = `🔍 Resultados para: "${rawText.trim()}"\n\n`;

    for (let i = 1; i < rows.length; i++) {
      const fila = rows[i];
      if (!fila || fila.length < 4) continue;

      const tituloOriginal = String(fila[3] || ""); 
      const autorOriginal  = String(fila[4] || "Desconocido");
      const isbnOriginal   = String(fila[5] || "-");
      const tituloLimpio = normalizar(tituloOriginal);

      if (tituloLimpio.includes(searchText)) {
        respuestaTexto += `📖 Obra: ${tituloOriginal}\n✍️ Autor: ${autorOriginal}\n🔢 ISBN: ${isbnOriginal}\n\n---\n`;
        encontrado = true;
      }
    }

    if (!encontrado) {
      respuestaTexto = `❌ No encontré coincidencias para "${rawText}".`;
    }

    // --- LOG PARA VALIDAR EN TERMINAL ---
    console.log("=== RESPUESTA GENERADA ===");
    console.log(respuestaTexto);
    console.log("==========================");

    // Envío a Teams (dará error en curl, pero es normal)
    await context.send(respuestaTexto).catch(err => {
        console.log("Nota: Envío a Teams falló (ID no encontrado), pero la búsqueda fue exitosa.");
    });

  } catch (error: any) {
    console.error("ERROR EN EL PROCESO:", error.message);
  }
});

export default app;