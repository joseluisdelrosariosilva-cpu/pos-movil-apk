import { Router } from "express";
import { getResumenDia } from "../controllers/resumen.controller.js";

const router = Router();

// GET /api/resumen - Resumen de ventas del día
router.get("/resumen", getResumenDia);

export default router;