import express from "express";
import googleApp from "./googleAppRoutes.js";

const routes = (app) => {
    app.route('/').get((req, res) => {
        res.status(200).send({titulo: "API - Google Apps with NodeJs"})
    })

    app.use(
        express.json(),
        googleApp,
    )
}

export default routes;