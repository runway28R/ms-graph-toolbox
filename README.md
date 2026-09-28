# Microsoft Graph Toolbox

A lightweight Python wrapper around Microsoft Graph for sending email, retrieving users, and working with SharePoint files.

## Features

- Authenticate with Microsoft Graph using Microsoft Entra ID application credentials
- Send email messages
- Retrieve users from Microsoft Graph
- List SharePoint sites and document libraries
- Browse folders and upload files to SharePoint

## Requirements

- Python 3.12 or newer
- A Microsoft Entra ID application registration
- Microsoft Graph API permissions
- A Microsoft 365 account with access to the required resources

## Installation

Install the package from PyPI:

```bash
pip install ms-graph-toolbox
```

## Authentication

The package uses application credentials:

- Client ID
- Client secret
- Tenant ID

Make sure the application has the required Microsoft Graph permissions and that administrator consent has been granted where necessary.

## Basic Usage

```python
import logging
import os

from ms_graph_toolbox.ms_graph_toolbox import ms_graph_toolbox

logger = logging.getLogger(__name__)

graph = ms_graph_toolbox(
    client_id=os.environ["MS_GRAPH_CLIENT_ID"],
    client_secret=os.environ["MS_GRAPH_CLIENT_SECRET"],
    tenant_id=os.environ["MS_GRAPH_TENANT_ID"],
    logger=logger,
)
```

The object automatically obtains an access token that is used by the other package functions.

## Sending Email

```python
from ms_graph_toolbox.graph_email import send_email

send_email(
    gph_object=graph,
    subject="Test message",
    content_type="Text",
    body="This message was sent using Microsoft Graph.",
    sender="sender@example.com",
    to_field="recipient@example.com",
)
```

Additional options include carbon-copy recipients, blind-carbon-copy recipients, message priority, and file attachments.

Multiple recipients can be provided as a comma-separated string:

```python
send_email(
    gph_object=graph,
    subject="Team update",
    content_type="Text",
    body="This message was sent to multiple recipients.",
    sender="sender@example.com",
    to_field="first@example.com, second@example.com",
)
```

## Retrieving Users

```python
from ms_graph_toolbox.graph_users import get_users

users = get_users(graph)

for user in users:
    print(user)
```

You can also filter the results by name, title, email address, alias, company, or selected fields:

```python
users = get_users(
    graph,
    select_data=["displayName", "mail", "jobTitle"],
    search_name="Alex",
)
```

## SharePoint Operations

SharePoint functionality is available through the `graph_sharepoint` class:

```python
from ms_graph_toolbox.graph_sharepoint import graph_sharepoint

sharepoint = graph_sharepoint(
    access_token=graph.access_token,
    logger=logger,
)
```

The SharePoint tools support:

- Finding a SharePoint site
- Listing document libraries
- Browsing folder contents
- Uploading files

See the [examples](https://github.com/runway28R/ms-graph-toolbox/tree/main/examples) directory for complete SharePoint usage examples.

## Examples

The repository contains complete examples for:

- Sending email
- Retrieving users
- Uploading files to SharePoint

The examples are available in the [examples](https://github.com/runway28R/ms-graph-toolbox/tree/main/examples) directory.

## Configuration

For local development, set these environment variables before running the examples.

### PowerShell

```powershell
$env:MS_GRAPH_CLIENT_ID = "your-client-id"
$env:MS_GRAPH_CLIENT_SECRET = "your-client-secret"
$env:MS_GRAPH_TENANT_ID = "your-tenant-id"
```

### Bash

```bash
export MS_GRAPH_CLIENT_ID="your-client-id"
export MS_GRAPH_CLIENT_SECRET="your-client-secret"
export MS_GRAPH_TENANT_ID="your-tenant-id"
```

Do not commit client secrets or other credentials to the repository.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
