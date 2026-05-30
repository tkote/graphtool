import argparse
import csv
import msal
import requests
import json
import os
import sys

client_id = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Azure CLI
authority = "https://login.microsoftonline.com/common"
scopes = ["User.Read.All"]
TOKEN_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")


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


def get_user(token, user_id):
    headers = {"Authorization": "Bearer " + token}
    url = (
        "https://graph.microsoft.com/v1.0/me"
        if user_id == "me"
        else f"https://graph.microsoft.com/v1.0/users/{user_id}"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()


def get_manager(token, user_id):
    headers = {"Authorization": "Bearer " + token}
    url = (
        "https://graph.microsoft.com/v1.0/me/manager"
        if user_id == "me"
        else f"https://graph.microsoft.com/v1.0/users/{user_id}/manager"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()


def get_direct_reports(token, user_id):
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


def search_users(token, query):
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


def get_group_members(token, group_ref, expand=False):
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


def _expand_members(token, members, seen_groups):
    users = []
    for m in members:
        if m.get("@odata.type") == "#microsoft.graph.group":
            gid = m["id"]
            if gid in seen_groups:
                continue
            seen_groups.add(gid)
            print(f"  Expanding group: {m.get('displayName', gid)}", file=sys.stderr)
            sub = _fetch_members_by_id(token, gid)
            users.extend(_expand_members(token, sub, seen_groups))
        else:
            users.append(m)
    return users


def search_groups(token, query):
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


def slim_group(group):
    return {
        "displayName": group.get("displayName", ""),
        "description": group.get("description", ""),
        "mail": group.get("mail", ""),
        "groupTypes": group.get("groupTypes", []),
    }


def print_group(group):
    print(group.get("displayName", "?") + f" <{group.get('mail', '')}>")
    if group.get("description"):
        print(f"  Description: {group.get('description')}")
    types = group.get("groupTypes", [])
    kind = "Microsoft 365" if "Unified" in types else ("Security" if not group.get("mailEnabled") else "Mail-enabled Security")
    print(f"  Type: {kind}")


def build_tree(token, user_id):
    user = get_user(token, user_id)
    display_name = user.get("displayName", user_id)
    print(f"Fetching: {display_name}", file=sys.stderr)
    reports = get_direct_reports(token, user.get("id", user_id))
    user["directReports"] = [build_tree(token, r["id"]) for r in reports]
    return user


def slim_user(node):
    return {
        "userPrincipalName": node.get("userPrincipalName", ""),
        "displayName": node.get("displayName", ""),
        "jobTitle": node.get("jobTitle", ""),
        "mail": node.get("mail", ""),
    }


def slim_tree(node):
    return {
        "userPrincipalName": node.get("userPrincipalName", ""),
        "displayName": node.get("displayName", ""),
        "directReports": [slim_tree(r) for r in node.get("directReports", [])],
    }


def print_user(node):
    print(node.get("displayName", "?") + f" <{node.get('mail', '')}>")
    if node.get("jobTitle"):
        print(f"  Title: {node.get('jobTitle')}")
    if node.get("userPrincipalName"):
        print(f"  UPN:   {node.get('userPrincipalName')}")


def flatten_tree(node, manager_upn=""):
    rows = []
    rows.append({
        "displayName": node.get("displayName", ""),
        "userPrincipalName": node.get("userPrincipalName", ""),
        "jobTitle": node.get("jobTitle", ""),
        "managerUPN": manager_upn,
    })
    for report in node.get("directReports", []):
        rows.extend(flatten_tree(report, manager_upn=node.get("userPrincipalName", "")))
    return rows


def print_tree(node, indent=0):
    prefix = "  " * indent + ("└─ " if indent > 0 else "")
    print(prefix + node.get("displayName", "?") + f" <{node.get('mail', '')}>")
    for report in node.get("directReports", []):
        print_tree(report, indent + 1)


def add_common_args(p):
    p.add_argument("user_id", nargs="?", default="me")
    p.add_argument("--format", choices=["text", "json"], default="text", dest="fmt")
    p.add_argument("--full", action="store_true", help="include all fields in JSON output")
    p.add_argument("--save", action="store_true", help="save result to a JSON file")


parser = argparse.ArgumentParser(description="Microsoft 365 graph tool")
subparsers = parser.add_subparsers(dest="command", required=True)

subparsers.add_parser("login", help="authenticate and cache token only")
add_common_args(subparsers.add_parser("self", help="show user info"))
add_common_args(subparsers.add_parser("manager", help="show manager info"))
p_tree = subparsers.add_parser("tree", help="show org tree")
p_tree.add_argument("user_id", nargs="?", default="me")
p_tree.add_argument("--format", choices=["text", "json", "csv"], default="text", dest="fmt")
p_tree.add_argument("--full", action="store_true", help="include all fields in JSON output")
p_tree.add_argument("--save", action="store_true", help="save result to a JSON file")
p_search = subparsers.add_parser("search", help="search users by name or email")
p_search.add_argument("query", help="partial name or email to search")
p_search.add_argument("--format", choices=["text", "json"], default="text", dest="fmt")
p_search.add_argument("--full", action="store_true", help="include all fields in JSON output")

p_search_groups = subparsers.add_parser("search-groups", help="search groups by name or email")
p_search_groups.add_argument("query", help="partial name or email to search")
p_search_groups.add_argument("--format", choices=["text", "json"], default="text", dest="fmt")
p_search_groups.add_argument("--full", action="store_true", help="include all fields in JSON output")

p_group_members = subparsers.add_parser("group-members", help="list members of a group")
p_group_members.add_argument("group_id", help="group display name")
p_group_members.add_argument("--format", choices=["text", "json", "csv"], default="text", dest="fmt")
p_group_members.add_argument("--full", action="store_true", help="include all fields in JSON output")
p_group_members.add_argument("--expand", action="store_true", help="recursively expand nested groups to individual users")

args = parser.parse_args()
token = get_token()

if args.command == "login":
    print("Token cached successfully.", file=sys.stderr)
    sys.exit(0)

elif args.command == "self":
    print(f"\nFetching user '{args.user_id}'...\n", file=sys.stderr)
    data = get_user(token, args.user_id)
    print("--- User ---", file=sys.stderr)
    if args.fmt == "json":
        print(json.dumps(data if args.full else slim_user(data), ensure_ascii=False, indent=2))
    else:
        print_user(data)
    if args.save:
        upn = data.get("userPrincipalName", args.user_id).replace("/", "_")
        output_file = f"user_{upn}.json"

elif args.command == "manager":
    print(f"\nFetching manager for '{args.user_id}'...\n", file=sys.stderr)
    data = get_manager(token, args.user_id)
    print("--- Manager ---", file=sys.stderr)
    if args.fmt == "json":
        print(json.dumps(data if args.full else slim_user(data), ensure_ascii=False, indent=2))
    else:
        print_user(data)
    if args.save:
        upn = args.user_id.replace("/", "_")
        output_file = f"manager_{upn}.json"

elif args.command == "tree":
    print(f"\nFetching org tree for '{args.user_id}'...\n", file=sys.stderr)
    data = build_tree(token, args.user_id)
    print("\n--- Org Tree ---", file=sys.stderr)
    if args.fmt == "json":
        print(json.dumps(data if args.full else slim_tree(data), ensure_ascii=False, indent=2))
    elif args.fmt == "csv":
        fields = ["displayName", "userPrincipalName", "jobTitle", "managerUPN"]
        writer = csv.DictWriter(sys.stdout, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        for row in flatten_tree(data):
            writer.writerow(row)
    else:
        print_tree(data)
    if args.save:
        upn = data.get("userPrincipalName", args.user_id).replace("/", "_")
        output_file = f"org_tree_{upn}.json"

elif args.command == "search":
    print(f"\nSearching users for '{args.query}'...\n", file=sys.stderr)
    results = search_users(token, args.query)
    print(f"--- {len(results)} result(s) ---", file=sys.stderr)
    if args.fmt == "json":
        out = results if args.full else [slim_user(u) for u in results]
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        for u in results:
            print_user(u)
            print()

elif args.command == "group-members":
    print(f"\nFetching members of group '{args.group_id}'...\n", file=sys.stderr)
    results = get_group_members(token, args.group_id, expand=args.expand)
    print(f"--- {len(results)} member(s) ---", file=sys.stderr)
    if args.fmt == "json":
        out = results if args.full else [slim_user(u) for u in results]
        print(json.dumps(out, ensure_ascii=False, indent=2))
    elif args.fmt == "csv":
        fields = ["displayName", "mail", "jobTitle", "userPrincipalName"]
        writer = csv.DictWriter(sys.stdout, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for u in results:
            writer.writerow(slim_user(u))
    else:
        for u in results:
            print_user(u)
            print()

elif args.command == "search-groups":
    print(f"\nSearching groups for '{args.query}'...\n", file=sys.stderr)
    results = search_groups(token, args.query)
    print(f"--- {len(results)} result(s) ---", file=sys.stderr)
    if args.fmt == "json":
        out = results if args.full else [slim_group(g) for g in results]
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        for g in results:
            print_group(g)
            print()

if hasattr(args, "save") and args.save:
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {output_file}", file=sys.stderr)
