// ============================================
// CONTROLLER: Productos (POS) - Versión RANGO
// ============================================
// Endpoint: GET /productos
// Lee desde un rango de celdas (no tabla formal)
// ============================================

import { abrirExcel, obtenerHojaPorNombre } from "../utils/excelHelper.js";

const NOMBRE_HOJA = "Productos"; // Ajusta al nombre real de tu hoja

/**
 * Normaliza nombres de campo
 */
const normalizarCampo = (texto) => {
  if (!texto) return "campo_" + Date.now();

  return texto
    .toString()
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[áéíóúüñ]/g, (letra) => {
      const map = { á: "a", é: "e", í: "i", ó: "o", ú: "u", ü: "u", ñ: "n" };
      return map[letra] || letra;
    })
    .replace(/[^a-z0-9_]/g, "");
};

/**
 * Filtra filas que están completamente vacías
 */
const filaTieneDatos = (fila) => {
  return fila.some(
    (celda) => celda !== null && celda !== undefined && celda !== "",
  );
};

export const getProductos = async (req, res) => {
  try {
    // 1. Abrir Excel
    const { workbook } = await abrirExcel();

    // 2. Obtener la hoja de productos
    const hoja = obtenerHojaPorNombre(workbook, NOMBRE_HOJA);

    // 3. Obtener rango utilizado
    const rango = hoja.usedRange();
    if (!rango) {
      return res
        .status(404)
        .json({ error: "No hay datos en la hoja de productos" });
    }

    // 4. Obtener matriz de valores
    const datos = rango.value();

    // 5. Validar que hay al menos encabezados
    if (datos.length < 2) {
      return res.status(404).json({ error: "No hay productos cargados" });
    }

    // 6. Extraer encabezados (primera fila)
    const encabezados = datos[0];

    // 7. Filtrar filas vacías y mapear a objetos
    const productos = datos
      .slice(1) // Saltamos encabezados
      .filter((fila) => filaTieneDatos(fila)) // Solo filas con contenido
      .map((fila) => {
        const producto = {};

        encabezados.forEach((encabezado, i) => {
          const campo = normalizarCampo(encabezado);
          producto[campo] =
            fila[i] !== undefined && fila[i] !== "" ? fila[i] : null;
        });

        return producto;
      });

    // 8. Responder con JSON
    res.json({
      total: productos.length,
      productos: productos,
    });
  } catch (error) {
    console.error("❌ Error en getProductos:", error.message);
    res.status(500).json({
      error: "Error al obtener productos",
      detalle: error.message,
    });
  }
};

export const getRecetas = async (req, res) => {
  try {
    const { workbook } = await abrirExcel();

    const hoja = obtenerHojaPorNombre(workbook, "Recetas");

    const rango = hoja.usedRange();
    if (!rango) {
      return res
        .status(404)
        .json({ error: "No hay datos en la hoja de recetas" });
    }

    const datos = rango.value();

    if (datos.length < 2) {
      return res.status(404).json({ error: "No hay recetas cargadas" });
    }

    const encabezados = datos[0];

    const recetas = datos
      .slice(1)
      .filter((fila) => filaTieneDatos(fila))
      .map((fila) => {
        const receta = {};

        encabezados.forEach((encabezado, i) => {
          const campo = normalizarCampo(encabezado);
          receta[campo] =
            fila[i] !== undefined && fila[i] !== "" ? fila[i] : null;
        });

        return receta;
      });

    res.json({
      total: recetas.length,
      recetas: recetas,
    });
  } catch (error) {
    console.error("❌ Error en getRecetas:", error.message);
    res.status(500).json({
      error: "Error al obtener recetas",
      detalle: error.message,
    });
  }
};

// ============================================
// GET /api/ingredientes - Lista ingredientes de recetas
// Lee columnas F:I de la hoja "Recetas" en datos.xlsx
// ============================================
export const getIngredientes = async (req, res) => {
  try {
    const { workbook } = await abrirExcel();

    const hoja = obtenerHojaPorNombre(workbook, "Recetas");

    const rango = hoja.usedRange();
    if (!rango) {
      return res
        .status(404)
        .json({ error: "No hay datos en la hoja de recetas" });
    }

    const datos = rango.value();

    if (datos.length < 2) {
      return res.status(404).json({ error: "No hay ingredientes cargados" });
    }

    // Mapeo fijo: columna F (índice 5) = Ingredientes, G (6) = Cantidad,
    // H (7) = Unidad, I (8) = Receta
    const ingredientes = datos
      .slice(1)
      .filter((fila) => fila[5] !== null && fila[5] !== undefined && fila[5] !== "")
      .map((fila) => ({
        ingrediente: fila[5] !== undefined && fila[5] !== "" ? fila[5] : null,
        cantidad: fila[6] !== undefined && fila[6] !== "" ? fila[6] : null,
        unidad: fila[7] !== undefined && fila[7] !== "" ? fila[7] : null,
        receta: fila[8] !== undefined && fila[8] !== "" ? fila[8] : null,
      }));

    res.json({
      total: ingredientes.length,
      ingredientes: ingredientes,
    });
  } catch (error) {
    console.error("❌ Error en getIngredientes:", error.message);
    res.status(500).json({
      error: "Error al obtener ingredientes",
      detalle: error.message,
    });
  }
};
