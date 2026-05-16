import { Router } from "express";
import { registrarGastos } from "../controllers/gastos.controller.js";

const router = Router();

// POST /api/gastos - Registrar gastos desde APK
router.post("/gastos", registrarGastos);

export default router;
