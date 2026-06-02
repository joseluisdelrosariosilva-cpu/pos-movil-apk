import { Router } from "express";
import { getProductos, getRecetas, getIngredientes } from "../controllers/pos.controller.js";

const router = Router();

// GET /api/productos - Lista todos los productos
router.get("/productos", getProductos);

// GET /api/recetas - Lista todas las recetas
router.get("/recetas", getRecetas);

// GET /api/ingredientes - Lista ingredientes de las recetas
router.get("/ingredientes", getIngredientes);

export default router;
