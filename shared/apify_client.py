#!/usr/bin/env python3
"""Shared Apify API client for all automation scripts."""

import time
from datetime import datetime
from typing import Optional

import requests


def log(msg: str) -> None:
    """Print timestamped log message."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def apify_request(
    token: str,
    method: str,
    endpoint: str,
    json_data: Optional[dict] = None,
    params: Optional[dict] = None
) -> requests.Response:
    """
    Make request to Apify API.

    Args:
        token: Apify API token
        method: HTTP method (GET, POST, etc.)
        endpoint: API endpoint (e.g., "/acts/{actor_id}/runs")
        json_data: Optional JSON body
        params: Optional query parameters

    Returns:
        requests.Response object

    Raises:
        requests.HTTPError: If the request fails
    """
    url = f"https://api.apify.com/v2{endpoint}"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.request(method, url, headers=headers, json=json_data, params=params)
    response.raise_for_status()
    return response


def run_actor_and_wait(
    token: str,
    actor_id: str,
    input_data: Optional[dict] = None,
    timeout: int = 180,
    poll_interval: int = 3
) -> dict:
    """
    Run an Apify actor and wait for completion.

    Args:
        token: Apify API token
        actor_id: The Apify actor ID to run
        input_data: Optional input data for the actor
        timeout: Maximum time to wait in seconds (default: 180)
        poll_interval: Time between status checks in seconds (default: 3)

    Returns:
        dict: The run data from Apify API

    Raises:
        RuntimeError: If the actor fails
        TimeoutError: If the actor doesn't complete within timeout
    """
    log(f"Starting Apify actor: {actor_id}")

    resp = apify_request(
        token,
        "POST",
        f"/acts/{actor_id}/runs",
        json_data=input_data or {},
        params={"timeout": timeout}
    )
    run_data = resp.json()["data"]
    run_id = run_data["id"]
    log(f"Run started: {run_id}")

    start = time.time()
    while time.time() - start < timeout:
        resp = apify_request(token, "GET", f"/actor-runs/{run_id}")
        status = resp.json()["data"]["status"]

        if status == "SUCCEEDED":
            log(f"Run completed: {run_id}")
            return resp.json()["data"]
        if status in ("FAILED", "ABORTED", "TIMED-OUT"):
            raise RuntimeError(f"Actor failed with status: {status}")

        log(f"Status: {status}... waiting")
        time.sleep(poll_interval)

    raise TimeoutError("Timeout waiting for actor")


def get_dataset_items(token: str, dataset_id: str) -> list:
    """
    Fetch items from an Apify dataset.

    Args:
        token: Apify API token
        dataset_id: The dataset ID to fetch from

    Returns:
        list: Items from the dataset
    """
    resp = apify_request(token, "GET", f"/datasets/{dataset_id}/items")
    return resp.json()


def get_key_value_store_record(token: str, store_id: str, key: str) -> bytes:
    """
    Fetch a record from an Apify key-value store.

    Args:
        token: Apify API token
        store_id: The key-value store ID
        key: The record key

    Returns:
        bytes: The record content
    """
    resp = apify_request(token, "GET", f"/key-value-stores/{store_id}/records/{key}")
    return resp.content
