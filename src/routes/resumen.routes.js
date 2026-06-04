import { Router } from "express";
import { getResumen } from "../controllers/resumen.controller.js";

const router = Router();

// GET /api/resumen?desde=YYYY-MM-DD&hasta=YYYY-MM-DD - Resumen de ventas por rango
router.get("/resumen", getResumen);

export default router;