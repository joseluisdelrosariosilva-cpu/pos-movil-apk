import { Router } from "express";
import { registrarMermas } from "../controllers/mermas.controller.js";

const router = Router();

// POST /api/mermas - Registrar mermas desde APK
router.post("/mermas", registrarMermas);

export default router;
