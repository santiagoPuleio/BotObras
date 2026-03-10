process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"; // Salta el proxy corporativo

import { stripMentionsText } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import * as xlsx from "xlsx";
import * as path from "path";

const storage = new LocalStorage();

// Crear la app básica (sin dependencias de Azure/OpenAI)
const app = new App({
  storage,
});

// --- FUNCIÓN DE LIMPIEZA Y NORMALIZACIÓN (El "Cerebro" de la búsqueda) ---
const normalizar = (texto: string): string => {
  return texto
    .toLowerCase()
    .normalize("NFD")               // Descompone caracteres con acento
    .replace(/[\u0300-\u036f]/g, "") // Elimina los acentos
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "") // Elimina puntuación
    .replace(/\s+/g, " ")           // Normaliza espacios múltiples
    .trim();
};

app.on("message", async (context) => {
  const activity = context.activity;
  const rawText = stripMentionsText(activity);
  const searchText = normalizar(rawText);

  // Comandos básicos
  if (rawText.trim() === "/reset") {
    storage.delete(activity.conversation.id);
    await context.send("Estado de conversación reiniciado.");
    return;
  }

  try {
    // 1. CARGAR EL EXCEL LOCAL
    const filePath = path.join(process.cwd(), "ObrasPlantaTabla.xlsx");
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // 2. CONVERTIR A MATRIZ DE FILAS
    const rows: any[][] = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    let encontrado = false;
    let respuesta = `🔍 **Resultados para: "${rawText.trim()}"**\n\n---\n`;

    // 3. RECORRER FILAS (Loop de validación inteligente)
    // Empezamos en i = 1 para saltar encabezados
    for (let i = 1; i < rows.length; i++) {
      const fila = rows[i];
      
      // Validar que la fila tenga suficientes columnas (hasta la F que es índice 5)
      if (!fila || fila.length < 4) continue;

      const tituloOriginal = String(fila[3] || ""); // Columna D
      const autorOriginal  = String(fila[4] || "Desconocido"); // Columna E
      const isbnOriginal   = String(fila[5] || "-"); // Columna F

      const tituloLimpio = normalizar(tituloOriginal);

      // VALIDACIÓN DE BÚSQUEDA PARCIAL
      // Si el texto del usuario está dentro del título, o viceversa
      if (tituloLimpio.includes(searchText) || (searchText.length > 3 && tituloLimpio.includes(searchText))) {
        respuesta += `📖 **Obra:** ${tituloOriginal}\n✍️ **Autor:** ${autorOriginal}\n🔢 **ISBN:** ${isbnOriginal}\n\n---\n`;
        encontrado = true;
      }
    }

    if (encontrado) {
      await context.send(respuesta);
    } else {
      await context.send(`❌ No encontré coincidencias para **"${rawText.trim()}"**.\n\n*Tip: Intenta buscar solo una palabra clave del título.*`);
    }

  } catch (error) {
    console.error(error);
    await context.send(`⚠️ Error crítico: No se pudo leer el archivo Excel.\nDetalle: ${error.message}`);
  }
});

export default app;