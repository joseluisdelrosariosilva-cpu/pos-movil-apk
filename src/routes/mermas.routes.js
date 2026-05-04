import express from "express";
import { registrarMermas } from "../controllers/mermas.controller.js";

const router = express.Router();

router.post("/", registrarMermas);

export default router;
