import csv
import io
import json
import os
import sys

import msal
import requests
from mcp.server.fastmcp import FastMCP

client_id = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Azure CLI
authority = "https://login.microsoftonline.com/common"
scopes = ["User.Read.All"]
TOKEN_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")

mcp = FastMCP("graphtool")


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_token():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE).read())

    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        print(flow["message"], file=sys.stderr)
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    if "access_token" not in result:
        raise RuntimeError("Authentication failed: " + result.get("error_description", ""))

    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------

def _get_user(token, user_id):
    headers = {"Authorization": "Bearer " + token}
    url = (
        "https://graph.microsoft.com/v1.0/me"
        if user_id == "me"
        else f"https://graph.microsoft.com/v1.0/users/{user_id}"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()


def _get_manager(token, user_id):
    headers = {"Authorization": "Bearer " + token}
    url = (
        "https://graph.microsoft.com/v1.0/me/manager"
        if user_id == "me"
        else f"https://graph.microsoft.com/v1.0/users/{user_id}/manager"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()


def _get_direct_reports(token, user_id):
    headers = {"Authorization": "Bearer " + token}
    url = (
        "https://graph.microsoft.com/v1.0/me/directReports"
        if user_id == "me"
        else f"https://graph.microsoft.com/v1.0/users/{user_id}/directReports"
    )
    members = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        members.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return members


def _build_tree(token, user_id, max_depth=None, _depth=0):
    user = _get_user(token, user_id)
    if max_depth is not None and _depth >= max_depth:
        user["directReports"] = []
    else:
        reports = _get_direct_reports(token, user.get("id", user_id))
        user["directReports"] = [_build_tree(token, r["id"], max_depth, _depth + 1) for r in reports]
    return user


def _slim_tree(node):
    return {
        "userPrincipalName": node.get("userPrincipalName", ""),
        "displayName": node.get("displayName", ""),
        "directReports": [_slim_tree(r) for r in node.get("directReports", [])],
    }


def _filter_managers_only(node):
    filtered = [_filter_managers_only(r) for r in node.get("directReports", []) if r.get("directReports")]
    return {**node, "directReports": filtered}


def _search_users(token, query):
    headers = {
        "Authorization": "Bearer " + token,
        "ConsistencyLevel": "eventual",
    }
    params = {
        "$search": f'"displayName:{query}" OR "mail:{query}" OR "userPrincipalName:{query}"',
        "$top": "25",
    }
    users = []
    url = "https://graph.microsoft.com/v1.0/users"
    first = True
    while url:
        resp = requests.get(url, headers=headers, params=params if first else None)
        first = False
        resp.raise_for_status()
        data = resp.json()
        users.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return users


def _search_groups(token, query):
    headers = {
        "Authorization": "Bearer " + token,
        "ConsistencyLevel": "eventual",
    }
    params = {
        "$search": f'"displayName:{query}" OR "mail:{query}"',
        "$top": "25",
    }
    groups = []
    url = "https://graph.microsoft.com/v1.0/groups"
    first = True
    while url:
        resp = requests.get(url, headers=headers, params=params if first else None)
        first = False
        resp.raise_for_status()
        data = resp.json()
        groups.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return groups


def _fetch_members_by_id(token, group_id):
    headers = {"Authorization": "Bearer " + token}
    members = []
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members"
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        members.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return members


def _expand_members(token, members, seen_groups):
    users = []
    for m in members:
        if m.get("@odata.type") == "#microsoft.graph.group":
            gid = m["id"]
            if gid in seen_groups:
                continue
            seen_groups.add(gid)
            sub = _fetch_members_by_id(token, gid)
            users.extend(_expand_members(token, sub, seen_groups))
        else:
            users.append(m)
    return users


def _get_group_members(token, group_ref, expand=False):
    headers = {"Authorization": "Bearer " + token}
    is_email = "@" in group_ref
    filter_expr = f"mail eq '{group_ref}'" if is_email else f"displayName eq '{group_ref}'"
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/groups",
        headers=headers,
        params={"$filter": filter_expr, "$select": "id,displayName,mail"},
    )
    resp.raise_for_status()
    value = resp.json().get("value", [])
    if not value:
        raise RuntimeError(f"Group not found: {group_ref}")
    group_id = value[0]["id"]
    members = _fetch_members_by_id(token, group_id)
    if not expand:
        return members
    return _expand_members(token, members, seen_groups={group_id})


def _slim_user(u):
    return {
        "displayName": u.get("displayName", ""),
        "mail": u.get("mail", ""),
        "jobTitle": u.get("jobTitle", ""),
        "userPrincipalName": u.get("userPrincipalName", ""),
    }


def _slim_group(g):
    return {
        "displayName": g.get("displayName", ""),
        "description": g.get("description", ""),
        "mail": g.get("mail", ""),
        "groupTypes": g.get("groupTypes", []),
    }


# ---------------------------------------------------------------------------
# MCP tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_user(user_id: str = "me") -> str:
    """Get user information. Pass a UPN (user@example.com), object ID, or 'me'."""
    token = get_token()
    data = _get_user(token, user_id)
    return json.dumps(_slim_user(data), ensure_ascii=False, indent=2)


@mcp.tool()
def get_manager(user_id: str = "me") -> str:
    """Get the manager of a user. Pass a UPN, object ID, or 'me'."""
    token = get_token()
    data = _get_manager(token, user_id)
    return json.dumps(_slim_user(data), ensure_ascii=False, indent=2)


@mcp.tool()
def get_org_tree(user_id: str = "me", max_depth: int | None = None, managers_only: bool = False) -> str:
    """Get the org tree rooted at the given user as JSON.

    user_id: UPN, object ID, or 'me'.
    max_depth: how many levels to descend (None = unlimited).
    managers_only: if True, exclude leaf nodes (users without direct reports).
    """
    token = get_token()
    data = _build_tree(token, user_id, max_depth=max_depth)
    if managers_only:
        data = _filter_managers_only(data)
    return json.dumps(_slim_tree(data), ensure_ascii=False, indent=2)


@mcp.tool()
def search_users(query: str) -> str:
    """Search users by partial display name, email, or UPN. Returns up to 25 results."""
    token = get_token()
    results = _search_users(token, query)
    return json.dumps([_slim_user(u) for u in results], ensure_ascii=False, indent=2)


@mcp.tool()
def search_groups(query: str) -> str:
    """Search groups by partial display name or email. Returns up to 25 results."""
    token = get_token()
    results = _search_groups(token, query)
    return json.dumps([_slim_group(g) for g in results], ensure_ascii=False, indent=2)


@mcp.tool()
def get_group_members(group_ref: str, expand: bool = False) -> str:
    """List members of a group.

    group_ref: group display name or email address.
    expand: if True, recursively expand nested groups so only users are returned.
    Returns CSV with columns: displayName, mail, jobTitle, userPrincipalName.
    """
    token = get_token()
    members = _get_group_members(token, group_ref, expand=expand)
    buf = io.StringIO()
    fields = ["displayName", "mail", "jobTitle", "userPrincipalName"]
    writer = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    for u in members:
        writer.writerow(_slim_user(u))
    return buf.getvalue()


if __name__ == "__main__":
    mcp.run()
