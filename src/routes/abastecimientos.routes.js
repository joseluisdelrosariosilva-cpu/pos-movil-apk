import { Router } from "express";
import { sincronizarAbastecimientos } from "../controllers/abastecimientos.controller.js";

const router = Router();

// POST /api/abastecimientos - Sincronizar abastecimientos desde APK
router.post("/abastecimientos", sincronizarAbastecimientos);

export default router;
