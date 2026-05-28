import { Router } from "express";
import { crearVenta } from "../controllers/ventas.controller.js";

const router = Router();

// POST /api/ventas - Registrar una nueva venta (en tiempo real)
router.post("/ventas", crearVenta);

export default router;
