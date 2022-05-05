import express from "express";
import GoogleAppController from "../controllers/googleAppController.js";

const router = express.Router();

router
    .get("/googleApp", GoogleAppController.runGoogleApp)    


export default router;