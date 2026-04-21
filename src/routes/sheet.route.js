import { Router } from "express";
import { getAllSheetsData } from "../controllers/sheetsController";

const router = Router();

router.get('/getAllSheetsData', getAllSheetsData);

export default router;