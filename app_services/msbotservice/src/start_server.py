# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.
#
# Mirrors this team's own app_services/msbotservice/src/start_server.py
# (CloudAdapter + start_agent_process over aiohttp), extended with this
# project's own HTTP routes (/api/form, /api/sas, /api/submit, /api/status,
# /api/search) alongside the standard /api/messages bot endpoint.

from os import environ
from microsoft_agents.hosting.core import AgentApplication, AgentAuthConfiguration
from microsoft_agents.hosting.aiohttp import (
    start_agent_process,
    jwt_authorization_middleware,
    CloudAdapter,
)
from aiohttp.web import Request, Response, Application, run_app

import routes


def start_server(
    agent_application: AgentApplication, auth_configuration: AgentAuthConfiguration
):
    async def entry_point(req: Request) -> Response:
        agent: AgentApplication = req.app["agent_app"]
        adapter: CloudAdapter = req.app["adapter"]
        return await start_agent_process(
            req,
            agent,
            adapter,
        )

    APP = Application()
    APP.router.add_post("/api/messages", entry_point)
    APP.router.add_get("/api/form", routes.form)
    APP.router.add_get("/api/sas", routes.sas)
    APP.router.add_post("/api/submit", routes.submit)
    APP.router.add_get("/api/status", routes.status)
    APP.router.add_get("/api/search", routes.search)
    APP["agent_configuration"] = auth_configuration
    APP["agent_app"] = agent_application
    APP["adapter"] = agent_application.adapter

    try:
        run_app(APP, host="0.0.0.0", port=int(environ.get("PORT", 3978)))
    except Exception as error:
        raise error
