import { Router } from "express";
import { crearVenta, sincronizarVentas } from "../controllers/ventas.controller.js";

const router = Router();

// POST /api/ventas - Registrar una nueva venta (en tiempo real)
router.post("/ventas", crearVenta);

// POST /api/sync - Sincronizar ventas offline desde APK
router.post("/sync", sincronizarVentas);

export default router;