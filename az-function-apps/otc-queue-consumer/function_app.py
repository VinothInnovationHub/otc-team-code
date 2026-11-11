"""Azure Functions entrypoint — pure hosting plumbing. Every real decision
(the generate_otc call, document finalization, status updates, proactive
messaging) lives in agent.py; this file only wires the queue trigger to it.

This is the ONLY piece of the project hosted on Azure Functions when the
bot/web UI itself runs on App Service — background/queue-triggered work has
no first-party App-Service-native equivalent, so it stays here, matching this
team's own existing pattern of small, single-purpose Function Apps alongside
their App-Service-hosted bot (see their mi-extrct function app).
"""

import asyncio
import json
import logging

import azure.functions as func

import agent
from Constants import Defaults

logger = logging.getLogger("otc_queue_consumer.function_app")
app = func.FunctionApp()


@app.queue_trigger(arg_name="msg", queue_name=Defaults.STORAGE_QUEUE_NAME, connection="StorageConnection")
# IMPORTANT: "StorageConnection" is an identity-based connection, deliberately
# NOT "AzureWebJobsStorage" — the Functions runtime storage account (this
# Function App's own, auto-provisioned one) is a *different* account from the
# project's business-data storage (STORAGE_ACCOUNT_NAME), which is where the
# App Service app's agent.py::enqueue_request() actually sends messages. Using
# AzureWebJobsStorage here would point the trigger at the wrong account's
# queue and nothing would ever fire. Requires these app settings
# (identity-based connection, no connection string):
#   StorageConnection__accountName = <STORAGE_ACCOUNT_NAME>
#   StorageConnection__credential  = managedidentity
#   StorageConnection__clientId    = <this Function App's managed identity client id>
#   StorageConnection__queueServiceUri = https://<STORAGE_ACCOUNT_NAME>.queue.core.windows.net
def process_queue_message(msg: func.QueueMessage) -> None:
    body = json.loads(msg.get_body().decode("utf-8"))
    asyncio.run(agent.process_request(body))
