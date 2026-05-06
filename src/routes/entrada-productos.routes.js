import { Router } from "express";
import { sincronizarEntradaProductos } from "../controllers/entrada-productos.controller.js";

const router = Router();

// POST /api/entrada-productos - Sincronizar productos nuevos desde APK
router.post("/entrada-productos", sincronizarEntradaProductos);

export default router;
