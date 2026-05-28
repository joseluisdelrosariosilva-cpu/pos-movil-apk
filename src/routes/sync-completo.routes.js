import { Router } from "express";
import { sincronizarCompleto } from "../controllers/sync-completo.controller.js";

const router = Router();

// POST /api/sync/completo - Sincronizar TODO en un solo request batch
router.post("/sync/completo", sincronizarCompleto);

export default router;
