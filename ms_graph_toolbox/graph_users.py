import requests


def get_users(gph_object, 
              select_data: str | None = None,  
              search_name: str | None = None,
              search_title: str | None = None,
              search_email: str | None = None,
              search_alias: str | None = None,
              search_company: str | None = None
              ) -> list[dict] | None:
    """
    Search for users in Microsoft Graph based on provided criteria.

    Notes:
    - companyName is not reliably supported in server-side $filter for all tenants/APIs,
      so company filtering is applied client-side to avoid "unsupported filter" errors.
    - This function uses `startswith(...)` for server-side partial matching where possible
      and escapes single quotes in filter values.

    Args:
        gph_object: An initialized ms_graph_toolbox object with valid access_token and logger.
        select_data: Comma-separated string or a list of properties to return (e.g. "displayName,mail,jobTitle").
        search_name: Filter users by displayName (partial, startswith).
        search_title: Filter users by jobTitle (partial, startswith).
        search_email: Filter users by mail (partial, startswith).
        search_alias: Filter users by alias (mailNickname / userPrincipalName / proxyAddresses) (exact or startswith).
        search_company: Filter users by companyName (partial, case-insensitive). Applied client-side.

    Returns:
        List of user dicts matching criteria, or None on error.
    """
    try:
        def esc(val: str) -> str:
            return val.replace("'", "''")

        if not any([search_name, search_title, search_email, search_alias, search_company]):
            gph_object.logger.warning("No filters provided; retrieving all users may be slow in large organizations.")

        # Build server-side filters (exclude companyName to avoid unsupported-filter errors)
        server_filters = []
        if search_name:
            server_filters.append(f"startswith(displayName,'{esc(search_name)}')")
        if search_title:
            server_filters.append(f"startswith(jobTitle,'{esc(search_title)}')")
        if search_email:
            server_filters.append(f"startswith(mail,'{esc(search_email)}')")
        if search_alias:
            # Try matching common alias forms
            alias_escaped = esc(search_alias)
            server_filters.append(
                f"(startswith(userPrincipalName,'{alias_escaped}') or startswith(mailNickname,'{alias_escaped}') "
                f"or proxyAddresses/any(x:x eq 'smtp:{alias_escaped}') or proxyAddresses/any(x:x eq 'SMTP:{alias_escaped}'))"
            )

        endpoint = "https://graph.microsoft.com/v1.0/users"
        params = {}
        if server_filters:
            params["$filter"] = " and ".join(server_filters)
        if select_data:
            if isinstance(select_data, list):
                params["$select"] = ",".join(select_data)
            else:
                params["$select"] = select_data
        # Include count if desired (note: some endpoints require ConsistencyLevel header for $count)
        params["$count"] = "true"

        headers = {"Authorization": f"Bearer {gph_object.access_token}",
                "ConsistencyLevel": "eventual"}
        users_list = []

        # Use params on first request; if @odata.nextLink is returned, follow it directly (it already contains params)
        url = endpoint
        first = True
        while True:
            if first:
                resp = requests.get(url, headers=headers, params=params)
                first = False
            else:
                resp = requests.get(url, headers=headers)  # nextLink already has query
            if resp.status_code != 200:
                gph_object.logger.error(f"Failed to retrieve users: {resp.status_code} - {resp.text}")
                return None

            data = resp.json()
            users = data.get("value", [])
            users_list.extend(users)

            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
            url = next_link

        # Apply company filter client-side (case-insensitive, partial match) if requested
        if search_company:
            sc = search_company.lower()
            users_list = [u for u in users_list if sc in (u.get("companyName") or "").lower()]
        
        gph_object.logger.debug(f"Retrieved {len(users_list)} users")
        return users_list

    except Exception as e:
        # Log exception details
        try:
            gph_object.logger.error(f"Exception occurred while searching for users: {e}")
        except Exception:
            # Fallback if logger not available
            pass
        return None